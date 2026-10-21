bl_info = {
    "name": "Tesis Landmark System",
    "author": "Josefa",
    "version": (0, 9, 9),
    # Trazabilidad: NORMAL/S1 canónico; ERROR independiente; timestamp solo informativo
    "blender": (5, 0, 0),
    "category": "3D View",
    "description": "Sistema de landmarks cefalométricos con cálculo de grosor de tejido blando y CCU",
}

import bpy  # type: ignore
import gpu  # type: ignore
import os
import math
import time
import shutil
import logging
import datetime
import statistics
import random
import threading
import tempfile
import textwrap
import openpyxl  # type: ignore
import blf  # type: ignore
from gpu_extras.batch import batch_for_shader  # type: ignore
from openpyxl.styles import Font, PatternFill  # type: ignore
import bmesh  # type: ignore
from contextlib import contextmanager
from typing import Any, cast
from mathutils import Vector, Matrix  # type: ignore
from bpy_extras import view3d_utils  # type: ignore
from mathutils.bvhtree import BVHTree  # type: ignore
from bpy.types import Panel, Operator, PropertyGroup, UIList, AddonPreferences  # type: ignore
from bpy.props import ( # type: ignore
    StringProperty,
    CollectionProperty,
    IntProperty,
    BoolProperty,
    FloatProperty,
    PointerProperty,
    EnumProperty,
)  # type: ignore
from decimal import Decimal

# =============================================================================
# DESCRIPCIÓN GENERAL DEL FLUJO DEL ADDON
# =============================================================================
# Este archivo implementa un sistema de apoyo para medición antropométrica en
# Blender, orientado a la identificación de landmarks craneofaciales y a la
# estimación de espesor de tejido blando sobre mallas derivadas de tomografía.
#
# La lógica funcional se organiza en cinco grandes bloques:
# 1) carga de casos desde Excel;
# 2) importación y validación de mallas ósea y blanda;
# 3) marcado de landmarks y cálculo de grosor;
# 4) manejo del caso de cuello (CCU) y mediciones complementarias;
# 5) guardado estructurado de resultados en Excel.
#
# Importante: la identificación anatómica inicial es una decisión del observador.
# El programa automatiza la proyección geométrica y el cálculo del espesor, pero
# no reemplaza la selección anatómica hecha por el investigador.
#
# La diferencia metodológica clave es esta:
# - DECISIÓN HUMANA: punto de referencia anatómico, validación visual, revisión.
# - AUTOMATIZACIÓN: ray cast, BVH, cálculo de distancia entre superficies,
#   reconstrucción de objetos y exportación a Excel.
#
# El archivo conserva la trazabilidad metodológica e incorpora mejoras de
# rendimiento, reconstrucción CCU, control de superficies y una interfaz simplificada.
# =============================================================================

# =====================================================
# LOGGING
# =====================================================

# Evita que Python escriba .pyc/__pycache__ para este flujo de trabajo.
# Guarda logs en carpeta temporal del usuario, no dentro del addon.
_LOG_DIR = os.path.join(tempfile.gettempdir(), "tesisjo_logs")
os.makedirs(_LOG_DIR, exist_ok=True)
# Referencia al FileHandler para cerrarlo en unregister
_LOG_FH = None

log = logging.getLogger("tesisjo")
if not log.handlers:
    log.setLevel(logging.DEBUG)
    _LOG_FH = logging.FileHandler(os.path.join(_LOG_DIR, "tesisjo.log"), encoding="utf-8")
    _LOG_FH.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    log.addHandler(_LOG_FH)


@contextmanager
def perf_section(name, results=None):
    """Mide una fase importante sin emitir mensajes por frame."""
    started = time.perf_counter()
    try:
        yield
    finally:
        elapsed = time.perf_counter() - started
        if results is not None:
            results[name] = elapsed


def _print_perf_block(label, results):
    """Emite un resumen estable y compacto de una operación completa."""
    print(f"[TESIS PERF {label}]")
    for name, elapsed in results.items():
        print(f"{name}: {elapsed:.3f} s")


_EXCEL_IO_LOCK = threading.RLock()
_EXCEL_INDEX_CACHE = {}
_EXCEL_BACKUP_STAMPS = {}
_ACTIVE_SAVE_PERF = None


def _memory_diagnostic(case_key=""):
    """Diagnóstico barato de recursos addon sin purgar datos ajenos."""
    owned_meshes = [mesh for mesh in bpy.data.meshes if bool(mesh.get("tesis_owned"))]
    orphaned = [mesh for mesh in owned_meshes if mesh.users == 0]
    orphan_vertices = sum(len(mesh.vertices) for mesh in orphaned)
    orphan_polygons = sum(len(mesh.polygons) for mesh in orphaned)
    log.info(
        "Memoria caso=%s objects=%d meshes=%d orphan_meshes=%d orphan_vertices=%d orphan_polygons=%d",
        case_key, len(bpy.data.objects), len(bpy.data.meshes), len(orphaned),
        orphan_vertices, orphan_polygons,
    )


def _excel_cache_key(filepath, sheet_name):
    try:
        stat = os.stat(filepath)
        stamp = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        stamp = (None, None)
    return os.path.abspath(filepath), sheet_name, stamp


def _cached_sheet_index(filepath, ws, key_builder):
    """Devuelve un índice de filas y lo conserva mientras no cambie el XLSX."""
    cache_key = _excel_cache_key(filepath, ws.title)
    cached = _EXCEL_INDEX_CACHE.get(cache_key)
    if cached is not None:
        return cached
    index = {}
    for row in range(2, ws.max_row + 1):
        key = key_builder(row)
        if key:
            index[key] = row
    _EXCEL_INDEX_CACHE[cache_key] = index
    return index


def _invalidate_excel_indexes(filepath=None):
    if filepath is None:
        _EXCEL_INDEX_CACHE.clear()
        return
    absolute = os.path.abspath(filepath)
    for key in list(_EXCEL_INDEX_CACHE):
        if key[0] == absolute:
            del _EXCEL_INDEX_CACHE[key]


def _refresh_excel_index_stamps(filepath):
    """Reasigna índices existentes al sello nuevo tras un guardado propio."""
    try:
        stat = os.stat(filepath)
        new_stamp = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        return
    absolute = os.path.abspath(filepath)
    for key in list(_EXCEL_INDEX_CACHE):
        if key[0] != absolute:
            continue
        value = _EXCEL_INDEX_CACHE.pop(key)
        _EXCEL_INDEX_CACHE[(absolute, key[1], new_stamp)] = value



# Cache global de BVH Trees para ray-casting eficiente.
_BVH_CACHE = {
    "skull": None,
    "soft": None,
    "skull_name": "",
    "soft_name": "",
    "skull_sig": None,
    "soft_sig": None,
}

_SOFT_COMPONENT_CACHE = {
    "signature": None,
    "main_bvh": None,
    "ratio": 0.0,
    "face_count": 0,
    "total_faces": 0,
    "status": "AMBIGUOUS",
    "obj_name": "",
}

_SOFT_MAIN_COMPONENT_MIN_RATIO = 0.70


def invalidate_bvh_cache():
    """Limpia la caché BVH. Llamar al importar nuevas mallas o editar geometría."""
    global _SOFT_COMPONENT_CACHE
    _BVH_CACHE["skull"] = None
    _BVH_CACHE["soft"] = None
    _BVH_CACHE["skull_name"] = ""
    _BVH_CACHE["soft_name"] = ""
    _BVH_CACHE["skull_sig"] = None
    _BVH_CACHE["soft_sig"] = None
    _SOFT_COMPONENT_CACHE = {
        "signature": None,
        "main_bvh": None,
        "ratio": 0.0,
        "face_count": 0,
        "total_faces": 0,
        "status": "AMBIGUOUS",
        "obj_name": "",
    }


def _bvh_object_signature(obj):
    """Genera una firma barata del objeto para decidir si hay que reconstruir BVH."""
    if obj is None:
        return None
    try:
        data = obj.data
        mw = obj.matrix_world
        return (
            obj.name,
            getattr(data, "name", ""),
            len(getattr(data, "vertices", ()) or ()),
            len(getattr(data, "polygons", ()) or ()),
            tuple(round(float(v), 6) for row in mw for v in row),
        )
    except Exception:
        return (obj.name, None)


def _analyze_soft_component(soft_obj):
    """Analiza la componente conectada principal de la malla SOFT sin tocar la geometría original."""
    global _SOFT_COMPONENT_CACHE
    signature = _bvh_object_signature(soft_obj) if soft_obj is not None else None
    if (_SOFT_COMPONENT_CACHE.get("signature") == signature
            and _SOFT_COMPONENT_CACHE.get("obj_name") == getattr(soft_obj, "name", "")):
        return _SOFT_COMPONENT_CACHE
    if soft_obj is None:
        _SOFT_COMPONENT_CACHE = {
            "signature": None,
            "main_bvh": None,
            "ratio": 0.0,
            "face_count": 0,
            "total_faces": 0,
            "status": "AMBIGUOUS",
            "obj_name": "",
        }
        return _SOFT_COMPONENT_CACHE

    mesh = getattr(soft_obj, "data", None)
    if mesh is None:
        _SOFT_COMPONENT_CACHE = {
            "signature": None,
            "main_bvh": None,
            "ratio": 0.0,
            "face_count": 0,
            "total_faces": 0,
            "status": "AMBIGUOUS",
            "obj_name": soft_obj.name,
        }
        return _SOFT_COMPONENT_CACHE

    total_faces = len(getattr(mesh, "polygons", []) or [])
    if total_faces == 0:
        _SOFT_COMPONENT_CACHE = {
            "signature": None,
            "main_bvh": None,
            "ratio": 0.0,
            "face_count": 0,
            "total_faces": 0,
            "status": "AMBIGUOUS",
            "obj_name": soft_obj.name,
        }
        return _SOFT_COMPONENT_CACHE

    vertex_to_faces = {i: [] for i in range(len(getattr(mesh, "vertices", []) or []))}
    for fi, poly in enumerate(mesh.polygons):
        for v_idx in poly.vertices:
            vertex_to_faces.setdefault(v_idx, []).append(fi)

    visited_faces = set()
    component_faces = []
    for fi in range(total_faces):
        if fi in visited_faces:
            continue
        stack = [fi]
        comp = set()
        while stack:
            current = stack.pop()
            if current in visited_faces:
                continue
            visited_faces.add(current)
            comp.add(current)
            poly = mesh.polygons[current]
            for v_idx in poly.vertices:
                for other in vertex_to_faces.get(v_idx, []):
                    if other not in visited_faces:
                        stack.append(other)
        component_faces.append(comp)

    if not component_faces:
        ratio = 0.0
        face_count = 0
        main_bvh = None
        status = "AMBIGUOUS"
    else:
        largest = max(component_faces, key=lambda comp: len(comp))
        face_count = len(largest)
        ratio = face_count / float(total_faces)
        main_bvh = None
        status = "AMBIGUOUS"
        if ratio >= _SOFT_MAIN_COMPONENT_MIN_RATIO:
            depsgraph = bpy.context.evaluated_depsgraph_get() if hasattr(bpy.context, "evaluated_depsgraph_get") else None
            if depsgraph is not None:
                soft_eval = soft_obj.evaluated_get(depsgraph)
                mw = soft_eval.matrix_world.copy()
                verts_world = []
                face_index_map = {}
                polys = []
                for fi in sorted(largest):
                    poly = soft_eval.data.polygons[fi]
                    mapped = []
                    for v_idx in poly.vertices:
                        if v_idx not in face_index_map:
                            face_index_map[v_idx] = len(verts_world)
                            verts_world.append(mw @ soft_eval.data.vertices[v_idx].co)
                        mapped.append(face_index_map[v_idx])
                    polys.append(mapped)
                if verts_world and polys:
                    main_bvh = BVHTree.FromPolygons(verts_world, polys)
            if main_bvh is not None:
                status = "MAIN_COMPONENT"

    _SOFT_COMPONENT_CACHE = {
        "signature": signature,
        "main_bvh": main_bvh,
        "ratio": float(ratio),
        "face_count": int(face_count),
        "total_faces": int(total_faces),
        "status": status,
        "obj_name": soft_obj.name,
    }
    print(f"[TESIS SOFT]\ncomponentes: {len(component_faces)}\nprincipal: {ratio * 100.0:.1f}%\nfiltro externo: {'ACTIVO' if status == 'MAIN_COMPONENT' else 'NO APLICADO'}")
    return _SOFT_COMPONENT_CACHE


def prepare_soft_external_component(context, soft_obj):
    """Prepara la componente exterior principal del SOFT para medición y cuello."""
    global _SOFT_COMPONENT_CACHE
    if soft_obj is None:
        return None
    try:
        return _analyze_soft_component(soft_obj)
    except Exception:
        _SOFT_COMPONENT_CACHE = {
            "signature": None,
            "main_bvh": None,
            "ratio": 0.0,
            "face_count": 0,
            "total_faces": 0,
            "status": "AMBIGUOUS",
            "obj_name": soft_obj.name if soft_obj else "",
        }
        return _SOFT_COMPONENT_CACHE


def get_soft_measurement_bvh():
    """Devuelve el BVH activo para mediciones de tejido blando."""
    cache = _SOFT_COMPONENT_CACHE or {}
    if cache.get("signature") is None:
        soft_obj = bpy.data.objects.get(_BVH_CACHE.get("soft_name", ""))
        if soft_obj is not None:
            cache = prepare_soft_external_component(bpy.context, soft_obj) or cache
    if cache.get("main_bvh") is not None:
        return cache["main_bvh"]
    return _BVH_CACHE.get("soft")


def get_full_soft_bvh():
    """Devuelve exclusivamente el BVH SOFT completo ya preparado."""
    return _BVH_CACHE.get("soft")

def _ccu_ring_name(case=None, soft_obj=None):
    """Nombre estable del anillo cervical del caso activo."""
    if soft_obj is not None:
        return f"CCU_{soft_obj.name.replace('SOFT_', '')}"
    if case is not None:
        return f"CCU_{case.paciente}_{case.ojo}"
    return ""


def _get_current_ccu_ring(case=None, soft_obj=None):
    """Busca el anillo visual del caso sin crear ni recalcular datos."""
    name = _ccu_ring_name(case, soft_obj)
    return bpy.data.objects.get(name) if name else None


def _remove_current_ccu_ring(case=None, soft_obj=None):
    """Elimina solo la representación visual CCU; nunca modifica CCU_mm."""
    ring = _get_current_ccu_ring(case, soft_obj)
    if ring is not None:
        _safe_remove_object(ring)
        return True
    return False


def _compute_ring_length_mm(ring):
    """Longitud total del anillo en coordenadas WORLD, expresada en milímetros."""
    if ring is None or getattr(ring, "type", None) != 'MESH':
        return 0.0
    mesh = getattr(ring, "data", None)
    if mesh is None:
        return 0.0
    try:
        mw = ring.matrix_world
        total = 0.0
        for edge in mesh.edges:
            a_idx, b_idx = edge.vertices
            pa = mw @ mesh.vertices[a_idx].co
            pb = mw @ mesh.vertices[b_idx].co
            total += (pa - pb).length
        return float(total)
    except Exception:
        return 0.0


def _select_ccu_bvh(p_ant, p_post):
    """Elige la BVH más adecuada para CCU sin construir otra estructura.

    Se prefiere la componente principal del tejido blando para evitar placas.
    Si esa componente no contiene ambas referencias cervicales, se reutiliza la
    BVH SOFT completa ya cacheada. Esto evita recalcular geometría al entrar a CCU.
    """
    main_bvh = (_SOFT_COMPONENT_CACHE or {}).get("main_bvh")
    full_bvh = _BVH_CACHE.get("soft")
    if main_bvh is None:
        return full_bvh, "FULL_SOFT"

    try:
        ant_hit = main_bvh.find_nearest(Vector(p_ant))
        post_hit = main_bvh.find_nearest(Vector(p_post))
        ant_dist = float(ant_hit[3]) if ant_hit and ant_hit[0] is not None else float("inf")
        post_dist = float(post_hit[3]) if post_hit and post_hit[0] is not None else float("inf")
        if ant_dist <= 2.0 and post_dist <= 2.0:
            return main_bvh, "MAIN_COMPONENT"
    except Exception:
        pass

    return full_bvh, "FULL_SOFT_NECK_FALLBACK"


def _ccu_collect_radial_candidates(bvh, center, direction, min_radius, max_radius, scan_radius):
    """Recoge impactos sucesivos y conserva solo la banda radial cervical."""
    direction = direction.normalized()
    origin = center + direction * 0.05
    remaining = float(scan_radius)
    candidates = []
    last_radius = -1.0
    for _ in range(32):
        hit = bvh.ray_cast(origin, direction, remaining)
        point = hit[0]
        if point is None:
            break
        radius = (point - center).length
        if radius > last_radius + 0.10:
            if min_radius <= radius <= max_radius:
                candidates.append((float(radius), point.copy()))
            last_radius = radius
        travelled = (point - origin).length
        remaining -= max(0.05, travelled + 0.05)
        if remaining <= 0.0:
            break
        origin = point + direction * 0.05
    candidates.sort(key=lambda item: item[0])
    return candidates


def _ccu_solve_arc(indices, candidate_sets, directions, center, anchor_start, anchor_end,
                   reference_radius, jump_scale):
    """Selecciona una trayectoria radial suave y rellena huecos entre impactos válidos."""
    states = []
    for position, idx in enumerate(indices):
        if position == 0:
            states.append([(reference_radius, anchor_start.copy(), False)])
        elif position == len(indices) - 1:
            states.append([(reference_radius, anchor_end.copy(), False)])
        else:
            states.append([(radius, point.copy(), False) for radius, point in candidate_sets[idx]]
                          + [(None, None, True)])

    costs = [[float("inf")] * len(options) for options in states]
    parents = [[-1] * len(options) for options in states]
    costs[0][0] = 0.0
    for position in range(1, len(states)):
        for current_idx, (current_radius, _point, current_missing) in enumerate(states[position]):
            unary = (12.0 if current_missing else
                     0.08 * ((current_radius - reference_radius) /
                             max(reference_radius, 1.0)) ** 2)
            for previous_idx, (previous_radius, _previous_point, previous_missing) in enumerate(states[position - 1]):
                if costs[position - 1][previous_idx] == float("inf"):
                    continue
                if current_missing and previous_missing:
                    transition = 3.0
                elif current_missing or previous_missing:
                    transition = 8.0
                else:
                    delta = abs(current_radius - previous_radius)
                    transition = (delta / max(jump_scale, 1.0)) ** 2
                    if delta > jump_scale:
                        transition += 25.0 + (delta - jump_scale) * 2.0
                candidate_cost = costs[position - 1][previous_idx] + unary + transition
                if candidate_cost < costs[position][current_idx]:
                    costs[position][current_idx] = candidate_cost
                    parents[position][current_idx] = previous_idx

    current = min(range(len(states[-1])), key=lambda idx: costs[-1][idx])
    selected: list[Any] = [None] * len(indices)
    for position in range(len(indices) - 1, -1, -1):
        radius, point, missing = states[position][current]
        selected[position] = None if missing else (radius, point)
        current = parents[position][current] if position > 0 else -1

    interpolated = 0
    position = 0
    while position < len(selected):
        if selected[position] is not None:
            position += 1
            continue
        gap_start = position
        while position < len(selected) and selected[position] is None:
            position += 1
        left = gap_start - 1
        right = position
        if left < 0 or right >= len(selected):
            continue
        left_radius = selected[left][0]
        right_radius = selected[right][0]
        gap_length = right - left
        for offset, fill_position in enumerate(range(gap_start, right), start=1):
            fraction = offset / gap_length
            radius = left_radius * (1.0 - fraction) + right_radius * fraction
            idx = indices[fill_position]
            selected[fill_position] = (radius, center + directions[idx] * radius)
            interpolated += 1
    return selected, interpolated


def build_ccu_ring_geometry(context, case=None, state=None, soft_obj=None, *, reconstruction=False):
    """Construye o reconstruye el anillo visual de circunferencia cervical.

    DECISIÓN DEL OBSERVADOR:
    ``neck_ant`` y ``neck_post`` definen las referencias cervicales.

    AUTOMATIZACIÓN:
    a partir de esas referencias se busca la superficie cutánea alrededor del
    cuello usando la BVH ya cacheada. Si ``reconstruction=True`` el anillo es
    únicamente una representación visual de una CCU histórica: nunca modifica
    ``case.ccu_mm`` ni escribe en Excel.
    """
    if state is None:
        state = context.scene.tesis_state
    if case is None and state.cases and state.active_case < len(state.cases):
        case = state.cases[state.active_case]
    if case is None or not getattr(case, "has_neck", False):
        return None
    if soft_obj is None:
        _, soft_obj = obtener_mallas_paciente(quiet=True)
    if soft_obj is None:
        return None

    lm_ant = find_landmark_by_name(state.landmarks, "neck_ant") if getattr(state, "landmarks", None) else None
    lm_post = find_landmark_by_name(state.landmarks, "neck_post") if getattr(state, "landmarks", None) else None
    if lm_ant is None or lm_post is None or lm_ant.status != "SET" or lm_post.status != "SET":
        return None

    p_ant = Vector((lm_ant.x, lm_ant.y, lm_ant.z))
    p_post = Vector((lm_post.x, lm_post.y, lm_post.z))
    ap_vector = p_post - p_ant
    dist_ant_post = ap_vector.length
    if dist_ant_post < 1e-6:
        return None

    # METODOLOGÍA CCU:
    # La circunferencia cervical se calcula sobre la malla SOFT COMPLETA.
    # El filtro de componente principal se reserva para el automatismo facial;
    # aplicarlo al cuello puede excluir regiones cervicales válidas y deformar
    # el contorno. Reutilizamos la BVH completa cacheada, geométricamente
    # equivalente a reconstruirla desde el objeto SOFT evaluado.
    bvh = _BVH_CACHE.get("soft")
    if bvh is None:
        skull_obj, current_soft = obtener_mallas_paciente(quiet=True)
        if skull_obj is None or current_soft is None:
            return None
        ensure_bvh_cache(context, skull_obj, current_soft)
        bvh = _BVH_CACHE.get("soft")
    if bvh is None:
        return None
    bvh_source = "FULL_SOFT"

    tangent_ap = ap_vector.normalized()
    center = (p_ant + p_post) / 2.0

    up_world = Vector((0, 0, 1))
    plane_normal = up_world - tangent_ap * up_world.dot(tangent_ap)
    if plane_normal.length < 1e-6:
        plane_normal = Vector((1, 0, 0)) - tangent_ap * Vector((1, 0, 0)).dot(tangent_ap)
    if plane_normal.length < 1e-6:
        plane_normal = Vector((0, 1, 0)) - tangent_ap * Vector((0, 1, 0)).dot(tangent_ap)
    if plane_normal.length < 1e-6:
        return None
    plane_normal.normalize()

    bitangent = plane_normal.cross(tangent_ap)
    if bitangent.length < 1e-6:
        bitangent = Vector((0, 1, 0))
    bitangent.normalize()

    ring_name = _ccu_ring_name(case, soft_obj)
    old_ring = bpy.data.objects.get(ring_name)
    if old_ring is not None:
        _safe_remove_object(old_ring)

    segments = 256
    half_ap = dist_ant_post * 0.5
    min_radius = max(5.0, half_ap * 0.42)
    max_radius = max(half_ap + 12.0, half_ap * 1.70)
    scan_radius = max(max_radius * 1.25, half_ap * 2.25)
    jump_scale = max(5.0, half_ap * 0.18)

    directions = []
    candidate_sets = []
    for i in range(segments):
        angle = 2.0 * math.pi * i / segments
        direction = (
            math.cos(angle) * tangent_ap +
            math.sin(angle) * bitangent
        ).normalized()
        directions.append(direction)
        candidate_sets.append(
            _ccu_collect_radial_candidates(
                bvh, center, direction, min_radius, max_radius, scan_radius
            )
        )

    posterior_to_anterior_a = list(range(0, segments // 2 + 1))
    posterior_to_anterior_b = [0] + list(range(segments - 1, segments // 2, -1)) + [segments // 2]

    arc_a, interpolated_a = _ccu_solve_arc(
        posterior_to_anterior_a, candidate_sets, directions, center,
        p_post, p_ant, half_ap, jump_scale
    )
    arc_b, interpolated_b = _ccu_solve_arc(
        posterior_to_anterior_b, candidate_sets, directions, center,
        p_post, p_ant, half_ap, jump_scale
    )

    verts = [None] * segments
    for indices, selected in (
        (posterior_to_anterior_a, arc_a),
        (posterior_to_anterior_b, arc_b),
    ):
        for idx, choice in zip(indices, selected):
            if choice is not None:
                _radius, point = choice
                verts[idx] = point.copy()

    # El contorno pasa exactamente por las dos referencias cervicales.
    verts[0] = p_post.copy()
    verts[segments // 2] = p_ant.copy()

    unresolved = [idx for idx, point in enumerate(verts) if point is None]
    if unresolved:
        log.warning("CCU %s: %d/256 sectores no pudieron resolverse.", case.paciente, len(unresolved))
        return None

    mesh = bpy.data.meshes.new(ring_name)
    mesh["tesis_owned"] = True
    mesh["tesis_role"] = "CCU"
    ring = bpy.data.objects.new(ring_name, mesh)
    cols = ensure_tesis_collections()
    move_object_to_collection(ring, cols["ccu"])
    edges = [(i, (i + 1) % segments) for i in range(segments)]
    mesh.from_pydata(verts, edges, [])
    mesh.update()

    # La curva es una representación de control, no una medición nueva por sí sola.
    ring.display_type = 'WIRE'
    ring.hide_select = False
    ring.hide_viewport = False
    ring.show_in_front = True
    ring.color = (0.15, 0.85, 0.30, 1.0)

    interpolated_count = int(interpolated_a + interpolated_b)
    direct_count = int(segments - interpolated_count)
    visual_length = _compute_ring_length_mm(ring)

    ring["ccu_method"] = "CENTER_OUT_MULTIHIT_ANCHORED"
    ring["ccu_ignores_ct_cap"] = True
    ring["ccu_bvh_source"] = bvh_source
    ring["ccu_reference_radius_mm"] = float(half_ap)
    ring["ccu_min_radius_mm"] = float(min_radius)
    ring["ccu_max_radius_mm"] = float(max_radius)
    ring["ccu_reconstructed"] = bool(reconstruction)
    ring["ccu_interpolated_points"] = interpolated_count
    ring["ccu_direct_points"] = direct_count
    ring["ccu_quality"] = "DIRECT" if interpolated_count == 0 else "ESTIMATED"
    ring["ccu_visual_length_mm"] = float(visual_length)
    ring["ccu_saved_length_mm"] = float(getattr(case, "ccu_mm", 0.0))
    ring["neck_ant_xyz"] = tuple(float(v) for v in p_ant)
    ring["neck_post_xyz"] = tuple(float(v) for v in p_post)

    if reconstruction:
        saved_length = float(getattr(case, "ccu_mm", 0.0))
        if saved_length > 0:
            diff = abs(float(visual_length) - saved_length)
            if diff > 5.0:
                log.warning(
                    "Revisar CCU reconstruida de %s: visual %.2f / guardada %.2f mm (Δ %.2f).",
                    case.paciente, visual_length, saved_length, diff,
                )
            elif diff > 1.0:
                log.info("CCU reconstruida de %s difiere %.2f mm del valor guardado.", case.paciente, diff)

    return ring


# =====================================================
# ADDON PREFERENCES (ruta Excel configurable)
# =====================================================

_DEFAULT_EXCEL_PATH = os.environ.get(
    "FSTT_EXCEL_ADDON",
    os.path.join(os.path.expanduser("~"), "fstt_datos", "Excel_addon_tesis.xlsx"),
)
_ADDON_ID = (__package__ or __name__.split(".")[0] or "tesisjo")

class TesisPreferences(AddonPreferences):
    """
    Preferencias del addon Tesis Landmark System.
    Permite configurar la ruta del Excel y parámetros de sesión.
    """
    bl_idname = _ADDON_ID

    excel_path: StringProperty(
        name="Ruta Excel",
        description="Ruta al archivo Excel maestro de datos",
        default=_DEFAULT_EXCEL_PATH,
        subtype='FILE_PATH',
    )  # type: ignore

    disable_handler: BoolProperty(
        name="Desactivar handler de grosor (legacy)",
        description="Preferencia legacy sin efecto: la actualización por handler fue eliminada por estabilidad",
        default=False,
    )  # type: ignore

    observer_id: StringProperty(
        name="Observador",
        default="OBS1",
        description="Código del observador (se exporta a Mediciones_LM)",
    )  # type: ignore

    session_id: StringProperty(
        name="Sesión",
        default="S1",
        description="Sesión o tanda de medición (se exporta a Mediciones_LM)",
    )  # type: ignore

    replicate_id: IntProperty(
        name="Réplica",
        default=1,
        min=1,
        description="Número de repetición de la medición (se exporta a Mediciones_LM)",
    )  # type: ignore

    def draw(self, context: bpy.types.Context) -> None:
        """Dibuja el panel de preferencias en la UI de Blender."""
        layout = self.layout
        layout.prop(self, "excel_path")
        layout.prop(self, "disable_handler")
        layout.separator()
        layout.label(text="Sesión de medición (se exporta al Excel)")
        row = layout.row(align=True)
        row.prop(self, "observer_id")
        row.prop(self, "session_id")
        row.prop(self, "replicate_id")


def get_excel_path() -> str:
    """Retorna la ruta al Excel desde las preferencias del addon."""
    try:
        addon_entry = bpy.context.preferences.addons.get(_ADDON_ID)
        if addon_entry and addon_entry.preferences:
            return addon_entry.preferences.excel_path
    except Exception:
        pass
    return _DEFAULT_EXCEL_PATH


def _detect_excel_lock(filepath: str) -> bool:
    """Detecta si Excel/OneDrive tiene el archivo bloqueado (lock file ~$...)."""
    d, f = os.path.split(filepath)
    lock_name = "~$" + f
    return os.path.exists(os.path.join(d, lock_name))


def _open_workbook_safe(filepath: str, **kwargs):
    """Abre un workbook con retry automático si el archivo está bloqueado por OneDrive/Excel.
    Intenta hasta 3 veces con backoff de 1s.
    """
    if _detect_excel_lock(filepath):
        log.warning("Lock file detectado para %s — puede haber conflicto con Excel abierto.", filepath)

    with _EXCEL_IO_LOCK:
        started = time.perf_counter()
        last_exc = None
        for attempt in range(3):
            try:
                workbook = openpyxl.load_workbook(filepath, **kwargs)
                if _ACTIVE_SAVE_PERF is not None:
                    _ACTIVE_SAVE_PERF["abrir workbook"] = time.perf_counter() - started
                return workbook
            except PermissionError as e:
                last_exc = e
                log.warning("PermissionError al abrir Excel (intento %d/3): %s", attempt + 1, e)
                time.sleep(1.0 * (attempt + 1))
        raise PermissionError(f"No se pudo abrir el Excel tras 3 intentos: {last_exc}")


def _save_workbook_safe(wb, filepath: str):
    """Guarda con retry y conserva un backup por versión externa del archivo."""
    with _EXCEL_IO_LOCK:
        absolute = os.path.abspath(filepath)
        if os.path.exists(filepath):
            try:
                stat = os.stat(filepath)
                source_stamp = (stat.st_mtime_ns, stat.st_size)
            except OSError:
                source_stamp = None
            if _EXCEL_BACKUP_STAMPS.get(absolute) != source_stamp:
                try:
                    backup_started = time.perf_counter()
                    shutil.copy2(filepath, filepath + ".bak")
                    _EXCEL_BACKUP_STAMPS[absolute] = source_stamp
                    if _ACTIVE_SAVE_PERF is not None:
                        _ACTIVE_SAVE_PERF["crear backup"] = time.perf_counter() - backup_started
                except Exception as e:
                    log.warning("No se pudo crear backup: %s", e)

        last_exc = None
        for attempt in range(3):
            try:
                save_started = time.perf_counter()
                wb.save(filepath)
                if _ACTIVE_SAVE_PERF is not None:
                    _ACTIVE_SAVE_PERF["wb.save"] = time.perf_counter() - save_started
                invalidate_excel_cache()
                _refresh_excel_index_stamps(filepath)
                try:
                    stat = os.stat(filepath)
                    _EXCEL_BACKUP_STAMPS[absolute] = (stat.st_mtime_ns, stat.st_size)
                except OSError:
                    pass
                return
            except PermissionError as e:
                last_exc = e
                log.warning("PermissionError al guardar Excel (intento %d/3): %s", attempt + 1, e)
                time.sleep(1.0 * (attempt + 1))
        raise PermissionError(f"No se pudo guardar el Excel tras 3 intentos: {last_exc}")


# =====================================================
# HELPERS SEGUROS DE CONVERSIÓN
# =====================================================

def _safe_str(value, default="") -> str:
    """Convierte a str de forma segura, evitando None y valores vacíos raros."""
    if value is None:
        return default
    try:
        s = str(value).strip()
        if s.lower() in {"none", "nan"}:
            return default
        return s
    except Exception:
        return default


def _safe_float(value, default=0.0) -> float:
    """Convierte distintos tipos a float de forma robusta."""
    if value is None:
        return default

    if isinstance(value, bool):
        return float(value)

    if isinstance(value, (int, float, Decimal)):
        try:
            f = float(value)
            if math.isnan(f) or math.isinf(f):
                return default
            return f
        except Exception:
            return default

    if isinstance(value, str):
        s = value.strip().replace(",", ".")
        if not s or s.lower() in {"none", "nan", "null"}:
            return default
        try:
            f = float(s)
            if math.isnan(f) or math.isinf(f):
                return default
            return f
        except Exception:
            return default

    # Fórmulas, rich text, fechas, objetos raros de openpyxl, etc.
    try:
        f = float(value)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except Exception:
        return default


def _safe_int(value, default=0) -> int:
    try:
        return int(round(_safe_float(value, default)))
    except Exception:
        return default


def _case_pair_key(paciente, ojo):
    """Clave canónica paciente/ojo para todas las hojas del Excel.

    Excel representa una celda vacía como ``None``; CaseItem la guarda como
    cadena vacía. Normalizar ambos campos con ``_safe_str`` evita que un caso
    con Ojo vacío (p. ej., un identificador sin dato auxiliar) pierda su resumen al recargar.
    """
    return (_safe_str(paciente), _safe_str(ojo))


def _parse_age_years(value) -> int | None:
    """Extrae edad en años desde textos como '17', '17a' o '17 años'."""
    if value is None:
        return None

    txt = _safe_str(value).strip().lower()
    if not txt:
        return None

    digits = ""
    for ch in txt:
        if ch.isdigit():
            digits += ch
        elif digits:
            break

    if not digits:
        return None

    try:
        return int(digits)
    except Exception:
        return None


def _case_requires_review(case) -> bool:
    """Interpreta la columna legacy 'Ojo' como flag de atención/revisión."""
    return _safe_str(getattr(case, "ojo", "")).strip().lower() in {
        "1", "1.0", "true", "sí", "si", "yes"
    }


def _case_alert_label(case) -> str:
    return "Revisar" if _case_requires_review(case) else "Normal"


def _case_alert_int(case) -> int:
    return 1 if _case_requires_review(case) else 0


def _ws_set_value(ws, row: int, col: int, value: Any) -> None:
    """Asigna valor a una celda evitando falsos positivos de Pylance con openpyxl."""
    cast(Any, ws.cell(row, col)).value = value


@contextmanager
def open_workbook(filepath: str, save=False, **kwargs):
    """Context manager para abrir un workbook con cierre garantizado.
    Si save=True, guarda antes de cerrar (con retry+backup).
    """
    with _EXCEL_IO_LOCK:
        wb = _open_workbook_safe(filepath, **kwargs)
        try:
            yield wb
            if save:
                _save_workbook_safe(wb, filepath)
        finally:
            wb.close()


SHEET_CASES = "Casos"
SHEET_LANDMARKS = "Coordenadas"
SHEET_RESULTADOS = "Resultados"
SHEET_MEDICIONES = "Mediciones_LM"
SHEET_DISPONIBILIDAD = "Disponibilidad_LM"
SHEET_RESUMEN_LM = "Resumen_LM"

COL_EDAD = "Edad"
COL_GRUPO = "Grupo"
COL_SEXO = "Sexo"
COL_PACIENTE = "Paciente"
COL_PESO = "Peso"
COL_OJO = "Ojo"
COL_OBS = "Observaciones"
COL_CUELLO = "¿Tiene cuello? (1 = Sí, 0 = No)"
COL_FALTAN = "¿Faltan landmarks? (1 = Sí, 0 = No)"
COL_MEJOR = "¿Es la mejor carpeta del paciente? (1 = Sí)"
COL_RUTA_BLANDO = "Ruta malla blando"
COL_RUTA_HUESO = "Ruta malla hueso"
COL_INCLUIR = "¿Incluir?"

COL_PROCESADO = "Procesado (1 = Sí, 0 = No)"
COL_FECHA = "Fecha procesado"

# Cohorte persistente para el estudio de error intra/interobservador.
COL_ERROR_ESTUDIO = "Error_Estudio"
COL_ERROR_COHORTE = "Error_Cohorte"
COL_ERROR_ORDEN = "Error_Orden"
COL_ERROR_ELEGIBLE = "Error_Elegible"
_DEFAULT_ERROR_COHORT = "ERR30_2026"
_DEFAULT_ERROR_SEED = 27082026

# La medición canónica/principal de la tesis se identifica por contexto, no por fecha.
_PRIMARY_MEASUREMENT_SESSION = "S1"

# Blender conserva punteros internos a los strings devueltos por callbacks de
# EnumProperty dinámicos. La documentación oficial advierte que, si Python no
# mantiene referencias vivas a esos strings, Blender puede comportarse de forma
# errática o incluso cerrarse. Este caché conserva de forma permanente los
# identificadores/nombres de cohortes vistos durante la sesión.
_ERROR_COHORT_ENUM_STRING_REFS = {}

def _stable_error_enum_string(value):
    s = _safe_str(value).strip()
    cached = _ERROR_COHORT_ENUM_STRING_REFS.get(s)
    if cached is None:
        _ERROR_COHORT_ENUM_STRING_REFS[s] = s
        cached = s
    return cached

# Índice liviano de progreso por (caso, observador, sesión, réplica, cohorte).
# Se reconstruye al cargar el Excel y se actualiza tras cada guardado de error.
_ERROR_PROGRESS_CACHE = {}

# =============================================================================
# DEFINICIÓN DEL PROTOCOLO ANATÓMICO DEL ESTUDIO
# =============================================================================
# Esta sección fija la lista completa de landmarks usados por la tesis.
# Cada entrada define:
# - landmark óseo (hard id);
# - landmark blando (soft id), cuando corresponde;
# - nombre anatómico y descripción clínica del punto.
#
# El programa usa este catálogo como referencia universal para:
# - ordenar el flujo de marcado;
# - construir contadores por caso;
# - normalizar nombres y exportar datos;
# - comparar cobertura y ausencia anatómica entre pacientes.
#
# Importante: no todos los puntos tienen par hueso-tejido blando. En el caso del
# cuello, el sistema trabaja con puntos de tejido blando sin un equivalente óseo.
# =============================================================================
# Lista de landmarks esperados (hueso ↔ tejido blando).
# Cada entrada es una tupla: (hard_id, soft_id, descripción).
# Para cuello: soft_id=None (solo punto en tejido blando, sin par).
LANDMARKS_ESPERADOS = [
    # ── Centrales ──────────────────────────────────────────────
    # 1
    ("g",       "g'",       "Glabella - punto anterior en la línea media del frontal."),
    # 2
    ("n",       "n'",       "Nasion - punto medio en la sutura naso-frontal."),
    # 3
    ("mn",      "mn'",      "Mid-nasal - punto en la sutura internasal a mitad de camino entre n y rhi."),
    # 4
    ("rhi",     "rhi'",     "Rhinion - punto en la línea media en el extremo libre inferior de la sutura internasal."),
    # 5
    ("sn",      "sn'",      "Subnasale - punto en la línea media justo debajo de la espina nasal anterior."),
    # 6
    ("mp",      "mp'",      "Mid-philtrum - punto medio entre la base de la espina nasal y el prostión en el borde anterior de los maxilares."),
    # 7
    ("ls",      "ls'",      "Labrale superius - punto en la línea media en el borde más anterior de la cresta alveolar superior de los maxilares."),
    # 8
    ("li",      "li'",      "Labrale inferius - punto en la línea media en el borde más anterior de la cresta alveolar inferior de los maxilares."),
    # 9
    ("mls",     "mls'",     "Mentolabial sulcus - punto más profundo en la línea media del surco superior a la eminencia mentoniana."),
    # 10
    ("pg",      "pg'",      "Pogonion - punto más anterior en la línea media en la eminencia mentoniana de la mandíbula."),
    # 11
    ("gn",      "gn'",      "Gnathion - punto medio entre los puntos más anterior (pg) e inferior (m) en el mentón óseo."),
    # 12
    ("m",       "m'",       "Menton - punto más inferior en la línea media en la sínfisis mental de la mandíbula."),

    # ── Bilaterales ────────────────────────────────────────────
    # 13-14
    ("mso-izq", "mso'-izq", "Mid-supraorbital izquierdo - punto en el borde supraorbital en el plano medio sagital de la órbita."),
    ("mso-der", "mso'-der", "Mid-supraorbital derecho - punto en el borde supraorbital en el plano medio sagital de la órbita."),
    # 15-16
    ("mio-izq", "mio'-izq", "Mid-infraorbital izquierdo - punto en el borde infraorbital en el plano medio sagital de la órbita."),
    ("mio-der", "mio'-der", "Mid-infraorbital derecho - punto en el borde infraorbital en el plano medio sagital de la órbita."),
    # 17-18
    ("acp-izq", "acp'-izq", "Alare curvature point izquierdo - punto 3 mm lateral al borde de la apertura nasal."),
    ("acp-der", "acp'-der", "Alare curvature point derecho - punto 3 mm lateral al borde de la apertura nasal."),
    # 19-20
    ("go-izq",  "go'-izq",  "Gonion izquierdo - punto en lateral del borde del ángulo mandibular."),
    ("go-der",  "go'-der",  "Gonion derecho - punto en lateral del borde del ángulo mandibular."),
    # 21-22
    ("zy-izq",  "zy'-izq",  "Zygion izquierdo - extensión más lateral de la superficie lateral del arco cigomático."),
    ("zy-der",  "zy'-der",  "Zygion derecho - extensión más lateral de la superficie lateral del arco cigomático."),
    # 23-24
    ("mr-izq",  "mr'-izq",  "Mid-ramus izquierdo - punto en el centro de la rama mandibular."),
    ("mr-der",  "mr'-der",  "Mid-ramus derecho - punto en el centro de la rama mandibular."),
    # 25-26
    ("mmb-izq", "mmb'-izq", "Mid-mandibular border izquierdo - punto en el borde inferior del cuerpo mandibular a mitad de camino entre pg y go."),
    ("mmb-der", "mmb'-der", "Mid-mandibular border derecho - punto en el borde inferior del cuerpo mandibular a mitad de camino entre pg y go."),

    # ── Cuello (solo tejido blando, sin par hard) ──────────────
    ("neck_ant",  None, "Neck anterior - punto en la línea media del cuello anterior bajo el mentón."),
    ("neck_post", None, "Neck posterior - punto posterior del cuello para medir perímetro cervical."),
]

_NECK_IDS = frozenset({"NECK_ANT", "NECK_POST"})

def _normalize_lm(name):
    """Normaliza nombre de landmark."""
    return str(name).strip().upper().replace("-", "_")

_NECK_INDICES = []
_OTHER_INDICES = []
for _i, (_lid, _soft, _lbl) in enumerate(LANDMARKS_ESPERADOS):
    if _normalize_lm(_lid) in _NECK_IDS:
        _NECK_INDICES.append(_i)
    else:
        _OTHER_INDICES.append(_i)

_NECK_INDEX_SET = frozenset(_NECK_INDICES)

_LM_NAME_TO_IDX = {_normalize_lm(lid): i for i, (lid, _, __) in enumerate(LANDMARKS_ESPERADOS)}

_CENTRAL_INDICES = _OTHER_INDICES[:12]
_BILATERAL_INDICES = _OTHER_INDICES[12:]

_BILATERAL_PAIR_MAP = {}
for _p in range(0, len(_BILATERAL_INDICES) - 1, 2):
    _a = _BILATERAL_INDICES[_p]
    _b = _BILATERAL_INDICES[_p + 1]
    _BILATERAL_PAIR_MAP[_a] = _b
    _BILATERAL_PAIR_MAP[_b] = _a

# ── Tolerancia visual por landmark bilateral (mm) ─────────────
# Basado de forma OPERATIVA en errores/variabilidad reportados en la revisión global 2023.
# No usar como "diagnóstico de asimetría", sino como aviso visual de revisión.
_LM_ASYM_TOL_MM = {
    "mso": 2.0,
    "mio": 2.1,
    "acp": 2.1,   # aproximación operativa usando ac/ac'
    "go": 6.0,
    "zy": 3.0,
    "mr": 4.3,
    "mmb": 4.1,
}

# Fallback si no hay tolerancia específica
_DEFAULT_ASYM_TOL_MM = 2.0


def _base_pair_label(lm_id: str) -> str:
    lid = str(lm_id).strip().lower()
    if lid.endswith("-izq"):
        return lid[:-4]
    if lid.endswith("-der"):
        return lid[:-4]
    return lid


def _pair_tol_mm_for_index(idx: int, state=None) -> float:
    if state is not None and getattr(state, "asym_mode", "LANDMARK") == "FIXED":
        return float(getattr(state, "asym_warn_mm", _DEFAULT_ASYM_TOL_MM))
    if idx < 0 or idx >= len(LANDMARKS_ESPERADOS):
        return _DEFAULT_ASYM_TOL_MM
    lm_id = LANDMARKS_ESPERADOS[idx][0]
    base = _base_pair_label(lm_id)
    return float(_LM_ASYM_TOL_MM.get(base, _DEFAULT_ASYM_TOL_MM))


def _iter_bilateral_pairs():
    seen = set()
    for idx in _BILATERAL_INDICES:
        if idx in seen:
            continue
        pair_idx = _BILATERAL_PAIR_MAP.get(idx)
        if pair_idx is None:
            continue
        seen.add(idx)
        seen.add(pair_idx)
        yield idx, pair_idx

# Orden secuencial de marcado (cuello excluido).
_BILATERAL_PAIRED = list(_BILATERAL_INDICES[:6])
_BILATERAL_GO_IZQ = list(_BILATERAL_INDICES[6::2])
_BILATERAL_GO_DER = list(_BILATERAL_INDICES[7::2])
_SEQUENTIAL_ORDER = (
    list(_CENTRAL_INDICES)
    + _BILATERAL_PAIRED
    + _BILATERAL_GO_IZQ
    + _BILATERAL_GO_DER
)
_SEQ_POS_OF = {idx: pos for pos, idx in enumerate(_SEQUENTIAL_ORDER)}

def _next_sequential_idx(current_idx):
    """Retorna el siguiente índice de landmark en el orden secuencial, o None."""
    pos = _SEQ_POS_OF.get(current_idx)
    if pos is None or pos + 1 >= len(_SEQUENTIAL_ORDER):
        return None
    return _SEQUENTIAL_ORDER[pos + 1]

def _prev_sequential_idx(current_idx):
    """Retorna el índice anterior de landmark en el orden secuencial, o None."""
    pos = _SEQ_POS_OF.get(current_idx)
    if pos is None or pos <= 0:
        return None
    return _SEQUENTIAL_ORDER[pos - 1]

# =====================================================
# DATA STRUCTURES
# =====================================================

_CASE_PROPERTY_UPDATE_SUPPRESS = False
_ACTIVE_CASE_CHANGE_SUPPRESS = False
_FILTER_SELECTION_SUPPRESS = False


def _mark_case_dirty(self, context):
    state = getattr(getattr(context, "scene", None), "tesis_state", None)
    if state is not None and not _CASE_PROPERTY_UPDATE_SUPPRESS:
        state.dirty = True


def _switch_active_case(context, new_index, *, sync=True, force_sync=False):
    """Cambia de caso de forma atómica usando ``committed_active_case`` como verdad de escena.

    ``active_case`` puede cambiar primero por la UIList; ``committed_active_case`` conserva
    el índice del paciente que realmente está sincronizado con landmarks/objetos.
    """
    global _ACTIVE_CASE_CHANGE_SUPPRESS
    state = context.scene.tesis_state if context and getattr(context, "scene", None) else None
    if state is None or not getattr(state, "cases", None):
        return False
    if new_index < 0 or new_index >= len(state.cases):
        return False

    committed = int(getattr(state, "committed_active_case", -1))
    old_index = committed if 0 <= committed < len(state.cases) else -1

    if getattr(state, "dirty", False) and old_index >= 0 and old_index != new_index:
        state.filter_notice = "Hay cambios sin guardar. Guarda o descarta antes de cambiar de caso."
        if state.active_case != old_index:
            _ACTIVE_CASE_CHANGE_SUPPRESS = True
            try:
                state.active_case = old_index
            finally:
                _ACTIVE_CASE_CHANGE_SUPPRESS = False
        return False

    # Si el mismo caso ya está comprometido y sincronizado, un filtro no debe
    # reabrir el Excel ni destruir/recrear su escena.
    if old_index == new_index and not force_sync:
        if state.active_case != new_index:
            _ACTIVE_CASE_CHANGE_SUPPRESS = True
            try:
                state.active_case = new_index
            finally:
                _ACTIVE_CASE_CHANGE_SUPPRESS = False
        state.filter_notice = ""
        return True

    if _marking_modal_active:
        _finish_marking_modal()

    # Limpiar siempre el caso REALMENTE comprometido, incluso si la UIList ya
    # cambió active_case al índice solicitado antes de disparar el callback.
    if old_index >= 0:
        _ACTIVE_CASE_CHANGE_SUPPRESS = True
        try:
            state.active_case = old_index
        finally:
            _ACTIVE_CASE_CHANGE_SUPPRESS = False
        _OBJ_NAME_TO_LM.clear()
        _clear_case_visual_objects()
        invalidate_mallas_cache()
        invalidate_bvh_cache()

    _ACTIVE_CASE_CHANGE_SUPPRESS = True
    try:
        state.previous_active_case = old_index if old_index >= 0 else new_index
        state.active_case = new_index
    finally:
        _ACTIVE_CASE_CHANGE_SUPPRESS = False

    if sync:
        sync_landmarks_state_from_excel(state.cases[new_index], state)
    elif not state.landmarks:
        _init_landmarks_for_case(state)

    state.committed_active_case = new_index
    state.dirty = False
    state.filter_notice = ""
    return True

class CaseItem(PropertyGroup):
    """Representa un caso clínico (paciente + ojo) con metadatos y rutas a mallas."""
    paciente: StringProperty(name="Paciente") # type: ignore
    ojo: StringProperty(
        name="Alerta (Excel: Ojo)",
        description="Flag legacy del Excel. 1 = caso que requiere atención/revisión; 0 = caso sin alerta"
    ) # type: ignore

    edad: StringProperty(name="Edad") # type: ignore
    sexo: StringProperty(name="Sexo") # type: ignore
    grupo: StringProperty(name="Grupo") # type: ignore
    peso: StringProperty(name="Peso") # type: ignore
    obs: StringProperty(name="Observaciones", description="Notas clínicas o comentarios sobre el caso", update=_mark_case_dirty) # type: ignore

    ruta_blando: StringProperty(name="Ruta malla blando", subtype='FILE_PATH', description="Ruta al archivo de malla PLY de tejido blando (SOFT)") # type: ignore
    ruta_hueso: StringProperty(name="Ruta malla hueso", subtype='FILE_PATH', description="Ruta al archivo de malla PLY de hueso (SKULL)") # type: ignore

    procesado: BoolProperty(default=False) # type: ignore
    has_neck: BoolProperty(default=False, name="Tiene cuello", update=_mark_case_dirty) # type: ignore
    faltan_landmarks: BoolProperty(default=False, name="Faltan landmarks") # type: ignore
    ccu_mm: FloatProperty(default=0.0) # type: ignore
    estado_nutricional: StringProperty(default="") # type: ignore

    # Copia canónica del valor principal. En modo ERROR, ccu_mm puede representar
    # temporalmente la CCU de la sesión actual sin tocar el dato basal del estudio.
    baseline_ccu_mm: FloatProperty(default=0.0) # type: ignore
    baseline_estado_nutricional: StringProperty(default="") # type: ignore

    # Cohorte persistente de reproducibilidad (guardada en hoja Casos).
    error_study: BoolProperty(default=False, name="Caso de error") # type: ignore
    error_cohort: StringProperty(default="") # type: ignore
    error_order: IntProperty(default=0, min=0) # type: ignore
    error_eligible: BoolProperty(default=False) # type: ignore

    summary_set: IntProperty(default=0) # type: ignore
    summary_missing: IntProperty(default=0) # type: ignore
    summary_pend: IntProperty(default=0) # type: ignore
    summary_n_valid: IntProperty(default=0) # type: ignore
    summary_avg: FloatProperty(default=0.0) # type: ignore
    summary_median: FloatProperty(default=0.0) # type: ignore
    summary_sd: FloatProperty(default=0.0) # type: ignore
    summary_min: FloatProperty(default=0.0) # type: ignore
    summary_max: FloatProperty(default=0.0) # type: ignore
    summary_loaded: BoolProperty(default=False) # type: ignore
    summary_neck_ant_status: StringProperty(default="OMIT") # type: ignore
    summary_neck_post_status: StringProperty(default="OMIT") # type: ignore
    reconstruction_warning: StringProperty(default="") # type: ignore
    coordinate_frame_valid: BoolProperty(default=True) # type: ignore
    coordinate_review_required: BoolProperty(default=False) # type: ignore
    coordinate_frame: StringProperty(default="CASE_FRAME_V1") # type: ignore
    import_orientation: EnumProperty(
        name="Orientación usada",
        items=[
            ("KEEP", "KEEP", "Malla en el marco canónico"),
            ("LEGACY", "LEGACY", "Orientación histórica X+90 / Y-90"),
        ],
        default="KEEP",
    ) # type: ignore


class LandmarkItem(PropertyGroup):
    """Un landmark cefalométrico con coordenadas hard/soft y grosor calculado.
    
    Estados posibles:
      - OMIT: Pendiente de marcar
      - SET: Marcado con coordenadas válidas
      - MISSING: No existe anatómicamente (no aplicable)
    """
    name: StringProperty() # type: ignore
    status: StringProperty(default="OMIT") # type: ignore

    # Coordenadas del punto en TEJIDO BLANDO (world space, mm)
    x: FloatProperty(default=0.0) # type: ignore
    y: FloatProperty(default=0.0) # type: ignore
    z: FloatProperty(default=0.0) # type: ignore

    # Coordenadas del punto en HUESO (world space, mm)
    hx: FloatProperty(default=0.0) # type: ignore
    hy: FloatProperty(default=0.0) # type: ignore
    hz: FloatProperty(default=0.0) # type: ignore

    # Grosor de tejido blando = distancia(hard, soft) en mm
    thickness: FloatProperty(default=0.0, min=0.0) # type: ignore

    # Referencias a objetos Blender (esferas + línea) para manipulación directa
    hard_obj_name: StringProperty(default="") # type: ignore
    soft_obj_name: StringProperty(default="") # type: ignore
    line_obj_name: StringProperty(default="") # type: ignore

def _on_active_case_changed(state, context):
    """Callback de UIList: protege el caso sincronizado y evita perder cambios."""
    global _ACTIVE_CASE_CHANGE_SUPPRESS
    if _ACTIVE_CASE_CHANGE_SUPPRESS:
        return

    requested_index = int(getattr(state, "active_case", 0))
    committed_index = int(getattr(state, "committed_active_case", -1))

    if requested_index == committed_index:
        return

    if getattr(state, "dirty", False) and 0 <= committed_index < len(state.cases):
        _ACTIVE_CASE_CHANGE_SUPPRESS = True
        try:
            state.active_case = committed_index
        finally:
            _ACTIVE_CASE_CHANGE_SUPPRESS = False
        state.filter_notice = "Hay cambios sin guardar. Guarda o descarta antes de cambiar de caso."
        return

    # active_case ya contiene el valor solicitado por Blender. El helper usa
    # committed_active_case para limpiar el paciente que realmente estaba en escena.
    _switch_active_case(context, requested_index, sync=True, force_sync=(committed_index < 0))


def _on_filter_changed(self, context):
    """Selecciona el primer caso visible sin perder cambios sin guardar."""
    if _FILTER_SELECTION_SUPPRESS or not getattr(self, "cases", None):
        return
    if getattr(self, "dirty", False):
        try:
            self.filter_notice = "Hay cambios sin guardar. Guarda o descarta antes de cambiar de caso."
        except Exception:
            pass
        return
    _select_first_visible_case(context, sync=True)


def _on_measurement_context_changed(self, context):
    """Cambia de contexto sin reutilizar landmarks de otra medición.

    En NORMAL, la sesión canónica siempre es S1. El cambio se hace por contexto
    explícito; nunca por fecha de guardado.
    """
    if _FILTER_SELECTION_SUPPRESS or not getattr(self, "cases", None):
        return
    if getattr(self, "dirty", False):
        self.filter_notice = (
            "Hay cambios sin guardar. Guarda o descarta antes de cambiar observador, sesión o modo."
        )
        return

    # Si el caso actual deja de pertenecer a los filtros de la sesión, saltar al
    # primer caso visible. Si sigue visible, resincronizarlo desde la fuente
    # correcta (canónica en NORMAL; Mediciones_LM en ERROR).
    current = int(getattr(self, "active_case", 0))
    if 0 <= current < len(self.cases) and _case_matches_active_filters(self, self.cases[current]):
        _switch_active_case(context, current, sync=True, force_sync=True)
    else:
        _select_first_visible_case(context, sync=True)


def _error_cohort_items(self, context):
    """Ofrece cohortes ERROR con strings de vida estable para Blender RNA."""
    cohort_values = sorted({
        _safe_str(getattr(case, "error_cohort", "")).strip()
        for case in getattr(self, "cases", [])
        if bool(getattr(case, "error_study", False))
        and _safe_str(getattr(case, "error_cohort", "")).strip()
    })
    # No leer ``self.error_cohort`` desde su propio callback de items:
    # Blender necesita esta misma lista para resolver el valor del enum y una
    # lectura recursiva aquí puede volver inestable el RNA.
    if not cohort_values:
        cohort_values = [_DEFAULT_ERROR_COHORT]

    description = _stable_error_enum_string("Cohorte ERROR disponible")
    return [
        (
            _stable_error_enum_string(cohort),
            _stable_error_enum_string(cohort),
            description,
        )
        for cohort in cohort_values
    ]


class TesisState(PropertyGroup):
    # Casos
    cases: CollectionProperty(type=CaseItem) # type: ignore
    previous_active_case: IntProperty(default=0) # type: ignore
    committed_active_case: IntProperty(default=-1, min=-1) # type: ignore
    active_case: IntProperty(default=0, update=lambda self, ctx: _on_active_case_changed(self, ctx)) # type: ignore

    # Contexto de medición -------------------------------------------------
    measurement_mode: EnumProperty(
        name="Modo",
        items=[
            ("NORMAL", "Medición principal", "Trabaja sobre Coordenadas/Resultados canónicos"),
            ("ERROR", "Estudio de error", "Trabaja en una cohorte independiente sin sobrescribir la medición principal"),
        ],
        default="NORMAL",
        update=_on_measurement_context_changed,
    ) # type: ignore
    error_cohort: EnumProperty(
        name="Cohorte",
        items=_error_cohort_items,
        default=0,
        description="Identificador persistente de la cohorte intra/interobservador",
        update=_on_measurement_context_changed,
    ) # type: ignore
    error_only_pending: BoolProperty(
        name="Solo pendientes de esta sesión",
        default=True,
        description="Oculta casos ya completos para observador/sesión/réplica actuales",
        update=_on_filter_changed,
    ) # type: ignore
    error_include_neck: BoolProperty(
        name="Incluir CCU",
        default=True,
        description="Exige y registra los dos puntos cervicales y la CCU en el estudio de error",
        update=_on_filter_changed,
    ) # type: ignore
    error_sample_size: IntProperty(name="N cohorte", default=30, min=4, max=200) # type: ignore
    error_seed: IntProperty(name="Semilla", default=_DEFAULT_ERROR_SEED, min=0) # type: ignore

    # Filtros
    search: StringProperty(name="Buscar", default="", description="Buscar paciente por nombre o ID", update=_on_filter_changed) # type: ignore
    only_unprocessed: BoolProperty(name="Solo no procesados", default=True, description="Mostrar únicamente los casos que aún no fueron procesados", update=_on_filter_changed) # type: ignore
    filter_has_neck: BoolProperty(name="Solo con cuello", default=False, description="Filtrar casos que indican disponibilidad de medida cervical (Cuello)", update=_on_filter_changed) # type: ignore
    filter_sex: EnumProperty(
        name="Sexo",
        update=_on_filter_changed,
        description="Filtra los casos según el sexo registrado en Excel",
        items=[
            ("ALL", "Todos", ""),
            ("M", "Masculino", ""),
            ("F", "Femenino", ""),
        ],
        default="ALL"
    ) # type: ignore
    filter_age_min: IntProperty(
        name="Edad mínima",
        update=_on_filter_changed,
        default=0,
        min=0,
        soft_max=30,
        description="Edad mínima en años (0 = sin filtro)",
    ) # type: ignore
    filter_age_max: IntProperty(
        name="Edad máxima",
        update=_on_filter_changed,
        default=0,
        min=0,
        soft_max=30,
        description="Edad máxima en años (0 = sin filtro)",
    ) # type: ignore
    filter_age_band: EnumProperty(
        name="Grupo etario",
        update=_on_filter_changed,
        description="Filtra por los grupos etarios usados en la tesis",
        items=[
            ("ALL", "Todos", "Sin filtro de grupo etario"),
            ("6_11", "6–11", "Casos de 6 a 11 años"),
            ("12_18", "12–18", "Casos de 12 a 18 años"),
        ],
        default="ALL",
    ) # type: ignore
    filter_neck: EnumProperty(
        name="Cuello",
        update=_on_filter_changed,
        description="Filtra según disponibilidad de región cervical para CCU",
        items=[
            ("ALL", "Todos", "Con y sin cuello"),
            ("YES", "Sí", "Solo casos con cuello"),
            ("NO", "No", "Solo casos sin cuello"),
        ],
        default="ALL",
    ) # type: ignore

    # Filtros adicionales
    only_processed: BoolProperty(name="Solo procesados", default=False, description="Mostrar únicamente los casos marcados como procesados en el Excel", update=_on_filter_changed) # type: ignore
    filter_has_obs: EnumProperty(
        name="Observaciones",
        update=_on_filter_changed,
        description="Filtra por observaciones del caso",
        items=[
            ("ALL", "Todos", "Mostrar todos los casos"),
            ("WITH", "Con observaciones", "Solo casos con observaciones"),
            ("WITHOUT", "Sin observaciones", "Solo casos sin observaciones"),
        ],
        default="ALL",
    ) # type: ignore
    filter_has_lm_excel: BoolProperty(
        name="Con LM en Excel",
        update=_on_filter_changed,
        default=False,
        description="Mostrar solo casos que ya tienen al menos un landmark guardado en la hoja Coordenadas"
    ) # type: ignore
    filter_eye: EnumProperty(
        name="Alerta",
        update=_on_filter_changed,
        description="Filtra según el indicador de revisión registrado en la columna Ojo",
        items=[
            ("ALL", "Todas", ""),
            ("1", "Sí", "Caso con observaciones o que requiere revisión"),
            ("0", "No", "Caso sin alerta especial"),
        ],
        default="ALL"
    ) # type: ignore

    # Landmarks del caso actual
    landmarks: CollectionProperty(type=LandmarkItem) # type: ignore
    active_landmark: IntProperty(default=0) # type: ignore

    lm_group: EnumProperty(
        name="Vista LM",
        items=[
            ("SEQ", "Sec.", "Solo landmarks del flujo secuencial"),
            ("ALL", "Todos", "Mostrar todos los landmarks"),
            ("NECK", "Cuello", "Mostrar landmarks de cuello"),
        ],
        default="SEQ",
    ) # type: ignore

    lm_filter: EnumProperty(
        name="Filtro LM",
        items=[
            ("PENDING", "Pend.", "Mostrar solo landmarks pendientes"),
            ("ALL", "Todos", "Mostrar todos los landmarks"),
        ],
        default="ALL",
        description="Filtro de visualización de landmarks"
    ) # type: ignore

    lm_only_pending: BoolProperty(
        name="Solo pendientes",
        default=True,
        description="Compatibilidad antigua: usar filtro Pend."
    ) # type: ignore

    dirty: BoolProperty(default=False) # type: ignore
    filter_notice: StringProperty(default="") # type: ignore
    show_lm_names: BoolProperty(
        name="Ver nombres LM",
        default=True,
        description="Mostrar etiquetas con los nombres de cada landmark en la vista 3D"
    ) # type: ignore

    lm_count_set: IntProperty(default=0) # type: ignore
    lm_count_missing: IntProperty(default=0) # type: ignore
    lm_count_pending: IntProperty(default=0) # type: ignore

    case_count_total: IntProperty(default=0) # type: ignore
    case_count_processed: IntProperty(default=0) # type: ignore
    case_count_pending: IntProperty(default=0) # type: ignore
    case_status: EnumProperty(
    name="Estado",
    update=_on_filter_changed,
    items=[
        ("ALL", "Todos", ""),
        ("PENDING", "Pendientes", ""),
        ("PROCESSED", "Completos", ""),
    ],
    default="PENDING",
)  # type: ignore

    observer_id: StringProperty(
        name="Observador",
        default="OBS1",
        description="Código del observador",
        update=_on_measurement_context_changed,
    ) # type: ignore

    session_id: StringProperty(
        name="Sesión",
        default="S1",
        description=("Sesión o tanda de medición. En modo NORMAL, la medición canónica se guarda únicamente como S1; "
                     "las sesiones de error deben usar modo ERROR."),
        update=_on_measurement_context_changed,
    ) # type: ignore

    replicate_id: IntProperty(
        name="Réplica",
        default=1,
        min=1,
        description="Número de repetición de la medición",
        update=_on_measurement_context_changed,
    ) # type: ignore

    asym_mode: EnumProperty(
        name="Aviso asimetría",
        items=[
            ("FIXED", "Fijo", "Usar un único umbral fijo para todos los pares"),
            ("LANDMARK", "Por LM", "Usar tolerancia específica según el landmark"),
        ],
        default="LANDMARK",
    ) # type: ignore

    asym_warn_mm: FloatProperty(
        name="Aviso fijo (mm)",
        default=2.0,
        min=0.0,
        description="Umbral fijo de diferencia bilateral para aviso visual"
    ) # type: ignore

    import_orientation: EnumProperty(
        name="Orientación de importación",
        description=(
            "Conservar mantiene la orientación almacenada en los PLY. "
            "Legado aplica y hornea X +90° y luego Y -90°, como el código antiguo"
        ),
        items=[
            ("KEEP", "Conservar PLY", "No rotar las mallas al importarlas"),
            ("LEGACY", "Legado X+90 / Y-90", "Aplicar la orientación usada por el código antiguo"),
        ],
        default="KEEP",
    ) # type: ignore

# =====================================================
# HELPERS
# =====================================================

def _update_lm_counts(state):
    """Actualiza los contadores cacheados de landmarks en TesisState.
    Llamar cada vez que cambia el status de algún landmark.
    Si el caso activo no tiene cuello, neck_ant/neck_post no cuentan como pendientes.
    """
    case = state.cases[state.active_case] if state.cases and state.active_case < len(state.cases) else None
    s = m = p = 0
    for i in _OTHER_INDICES:
        if i >= len(state.landmarks):
            continue
        lm = state.landmarks[i]
        if lm.status == "SET":
            s += 1
        elif lm.status == "MISSING":
            m += 1
        else:
            p += 1
    state.lm_count_set = s
    state.lm_count_missing = m
    state.lm_count_pending = p
    # Los summaries del caso representan Excel/último guardado; no se alteran
    # al inicializar la lista temporal de landmarks del caso activo.

def _update_case_counts(state):
    """Actualiza contadores cacheados de casos para evitar escaneo en cada redraw."""
    total = len(state.cases)
    processed = sum(1 for case in state.cases if _case_is_complete(case, state=state))
    state.case_count_total = total
    state.case_count_processed = processed
    state.case_count_pending = max(0, total - processed)


def _case_facial_counts(case, state=None):
    if state is not None and state.cases and state.cases[state.active_case] == case and state.landmarks:
        statuses = [state.landmarks[idx].status for idx in _OTHER_INDICES if idx < len(state.landmarks)]
        return statuses.count("SET"), statuses.count("MISSING"), statuses.count("OMIT")
    return (int(getattr(case, "summary_set", 0)),
            int(getattr(case, "summary_missing", 0)),
            int(getattr(case, "summary_pend", 0)))


def _case_neck_statuses(case, state=None):
    if state is not None and state.cases and state.cases[state.active_case] == case and state.landmarks:
        return tuple(state.landmarks[idx].status for idx in _NECK_INDICES if idx < len(state.landmarks))
    return (str(getattr(case, "summary_neck_ant_status", "OMIT")).upper(),
            str(getattr(case, "summary_neck_post_status", "OMIT")).upper())


def _case_is_facial_complete(case, state=None):
    set_count, missing_count, pending_count = _case_facial_counts(case, state)
    return set_count + missing_count == len(_OTHER_INDICES) and pending_count == 0


def _case_is_neck_complete(case, state=None):
    if not bool(getattr(case, "has_neck", False)):
        return True
    neck_statuses = _case_neck_statuses(case, state)
    return (len(neck_statuses) == 2 and all(status == "SET" for status in neck_statuses)
            and _safe_float(getattr(case, "ccu_mm", 0.0), 0.0) > 0)


def _case_is_complete(case, state=None):
    return _case_is_facial_complete(case, state) and _case_is_neck_complete(case, state)


def _case_finalize_blockers(state, case):
    """Retorna los motivos metodológicos que impiden finalizar."""
    if state.cases and state.cases[state.active_case] == case and state.landmarks:
        facial_pending = sum(
            1 for idx in _OTHER_INDICES
            if idx >= len(state.landmarks) or state.landmarks[idx].status == "OMIT"
        )
    else:
        facial_pending = int(getattr(case, "summary_pend", 0))
    blockers = []
    if facial_pending:
        blockers.append(f"Falta completar {facial_pending} landmarks faciales.")
    if (_safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() == "ERROR"
            and not bool(getattr(state, "error_include_neck", True))):
        return blockers
    if not bool(getattr(case, "has_neck", False)):
        return blockers
    statuses = _case_neck_statuses(case, state)
    if len(statuses) < 2 or statuses[0] != "SET":
        blockers.append("Falta marcar el punto anterior del cuello.")
    if len(statuses) < 2 or statuses[1] != "SET":
        blockers.append("Falta marcar el punto posterior del cuello.")
    if _safe_float(getattr(case, "ccu_mm", 0.0), 0.0) <= 0:
        blockers.append("Falta confirmar la CCU.")
    return blockers


def _select_first_visible_case(context, *, sync=True):
    """Selecciona el primer caso visible sin cambiar de caso si hay cambios."""
    state = context.scene.tesis_state
    if not state.cases or state.dirty:
        return False
    index = next((i for i, case in enumerate(state.cases)
                  if _case_matches_active_filters(state, case)), None)
    if index is None:
        return False
    return _switch_active_case(context, index, sync=sync)


def _reset_case_summary(case):
    """Limpia los datos resumidos por caso para evitar mostrar caché vieja.

    Se usa al limpiar o descartar el caso actual.
    """
    if case is None:
        return
    case.summary_set = 0
    case.summary_missing = 0
    case.summary_pend = 0
    case.summary_n_valid = 0
    case.summary_avg = 0.0
    case.summary_median = 0.0
    case.summary_sd = 0.0
    case.summary_min = 0.0
    case.summary_max = 0.0
    case.summary_loaded = False


def _thickness_summary_from_values(values):
    """Resume espesores válidos excluyendo MISSING y pendientes."""
    valid = [float(v) for v in values if v is not None]
    if not valid:
        return {
            "n_valid": 0,
            "avg": 0.0,
            "median": 0.0,
            "sd": 0.0,
            "min": 0.0,
            "max": 0.0,
        }

    return {
        "n_valid": len(valid),
        "avg": round(statistics.mean(valid), 3),
        "median": round(statistics.median(valid), 3),
        "sd": round(statistics.stdev(valid), 3) if len(valid) > 1 else 0.0,
        "min": round(min(valid), 3),
        "max": round(max(valid), 3),
    }


def _bilateral_delta_for_index(state, idx):
    """Retorna |izq - der| en mm para el landmark bilateral actual."""
    pair_idx = _BILATERAL_PAIR_MAP.get(idx)
    if pair_idx is None:
        return None

    if idx >= len(state.landmarks) or pair_idx >= len(state.landmarks):
        return None

    a = state.landmarks[idx]
    b = state.landmarks[pair_idx]

    if a.status != "SET" or b.status != "SET":
        return None

    return abs(float(a.thickness) - float(b.thickness))


# Umbral visual fijo para marcar diferencias bilaterales en rojo
_UI_BILAT_WARN_MM = 2.0

# Límites individuales por landmark (dejar vacío = solo usa Δ bilateral)
_LM_MAX_WARN_MM = {}


def _get_landmark_index_by_id(lm_id: str):
    lid = str(lm_id).strip().lower()
    for i, spec in enumerate(LANDMARKS_ESPERADOS):
        if str(spec[0]).strip().lower() == lid:
            return i
    return None


def _pair_warn_exceeded(state, idx, delta=None):
    """Aviso visual bilateral uniforme; no altera el delta exportado."""
    if idx not in _BILATERAL_PAIR_MAP:
        return False
    if delta is None:
        delta = _bilateral_delta_for_index(state, idx)
    return delta is not None and float(delta) > _UI_BILAT_WARN_MM


def _landmark_ui_warn(state, idx):
    """True si el landmark debe resaltarse visualmente (Δ bilateral o grosor excesivo)."""
    if idx < 0 or idx >= len(state.landmarks):
        return False
    lm = state.landmarks[idx]
    lm_id = LANDMARKS_ESPERADOS[idx][0] if idx < len(LANDMARKS_ESPERADOS) else lm.name
    delta = _bilateral_delta_for_index(state, idx)
    if _pair_warn_exceeded(state, idx, delta):
        return True
    lim = _LM_MAX_WARN_MM.get(lm_id)
    if lim is not None and lm.status == "SET" and float(lm.thickness) > float(lim):
        return True
    return False


def normalize_landmark_name(name):

    """Normaliza nombre de landmark: strip, uppercase, guiones a underscores."""
    return _safe_str(name).strip().upper().replace("-", "_")

def find_landmark_by_name(landmarks, target_name):
    """Busca landmark por nombre normalizado. Retorna landmark o None."""
    target = normalize_landmark_name(target_name)
    idx = _LM_NAME_TO_IDX.get(target)
    if idx is not None and idx < len(landmarks):
        return landmarks[idx]
    # Fallback: búsqueda lineal (por si hay nombres custom)
    for lm in landmarks:
        if normalize_landmark_name(getattr(lm, "name", "")) == target:
            return lm
    return None

def ensure_scene_units_mm():
    scn = bpy.context.scene
    u = scn.unit_settings
    u.system = 'METRIC'
    u.length_unit = 'MILLIMETERS'
    u.scale_length = 0.001  # 1BU = 1mm

def _safe_remove_object(obj):
    """Desvincula un objeto sin tocar materiales ni borrar su datablock.

    Importante para estabilidad de Undo/Redo: Blender puede conservar referencias
    internas a Object/Mesh/Curve en la pila de deshacer. Vaciar slots de material o
    borrar manualmente el datablock justo antes de un undo puede dejar referencias
    inválidas en el depsgraph. Los datos huérfanos se dejan para la limpieza normal
    de Blender y no se purgan durante el flujo de marcado.
    """
    if obj is None:
        return False
    try:
        name = obj.name
    except Exception:
        name = "?"
    try:
        bpy.data.objects.remove(obj, do_unlink=True)
        return True
    except ReferenceError:
        return False
    except Exception as exc:
        log.warning("_safe_remove_object: error removiendo %s: %s", name, exc)
        return False


def _cleanup_auxiliary_datablocks():
    """Retira solo meshes/curves auxiliares del addon que quedaron sin usuarios."""
    _cleanup_owned_orphans()


def _cleanup_owned_orphan_curves():
    """Libera solo curvas propias del addon sin usuarios."""
    if _undo_in_progress or _marking_modal_active:
        return 0
    removed = 0
    for curve in list(bpy.data.curves):
        try:
            if bool(curve.get("tesis_owned")) and curve.users == 0:
                bpy.data.curves.remove(curve)
                removed += 1
        except (ReferenceError, RuntimeError):
            continue
    if removed:
        log.debug("Curvas addon huérfanas liberadas: %d", removed)
    return removed


def _cleanup_owned_orphans():
    """Limpia recursos addon huérfanos respetando las protecciones de runtime."""
    if _undo_in_progress or _marking_modal_active:
        global _DEFERRED_CLEANUP_PENDING, _DEFERRED_CLEANUP_TIMER
        _DEFERRED_CLEANUP_PENDING = True
        if not _DEFERRED_CLEANUP_TIMER:
            _DEFERRED_CLEANUP_TIMER = True
            bpy.app.timers.register(_run_deferred_cleanup, first_interval=0.5)
        return 0
    return _cleanup_owned_orphan_meshes() + _cleanup_owned_orphan_curves()


def _cleanup_owned_orphan_meshes():
    """Libera únicamente meshes propiedad del addon que ya no tienen usuarios."""
    global _DEFERRED_CLEANUP_PENDING, _DEFERRED_CLEANUP_TIMER
    if _undo_in_progress or _marking_modal_active:
        _DEFERRED_CLEANUP_PENDING = True
        if not _DEFERRED_CLEANUP_TIMER:
            _DEFERRED_CLEANUP_TIMER = True
            bpy.app.timers.register(_run_deferred_cleanup, first_interval=0.5)
        return 0
    removed = 0
    for mesh in list(bpy.data.meshes):
        try:
            if (mesh.name not in _LM_SPHERE_MESH_NAMES.values()
                    and bool(mesh.get("tesis_owned")) and mesh.users == 0):
                bpy.data.meshes.remove(mesh)
                removed += 1
        except (ReferenceError, RuntimeError):
            continue
    if removed:
        log.debug("Meshes addon huérfanas liberadas: %d", removed)
    return removed


def _run_deferred_cleanup():
    """Callback de main thread para liberar datablocks cuando sea seguro."""
    global _DEFERRED_CLEANUP_PENDING, _DEFERRED_CLEANUP_TIMER
    if _undo_in_progress or _marking_modal_active:
        return 0.5
    _DEFERRED_CLEANUP_PENDING = False
    _DEFERRED_CLEANUP_TIMER = False
    _cleanup_owned_orphans()
    return None


def clear_objects_prefix(prefixes):
    objs = [obj for obj in bpy.data.objects
            if any(obj.name.startswith(p) for p in prefixes)]
    for obj in objs:
        _safe_remove_object(obj)

def clear_landmark_objects(idx: int):
    """Desvincula todos los objetos pertenecientes a un landmark.

    Reconoce la convención estable ``00_g_*``, la convención transitoria ``1.g_*``
    y los metadatos persistentes ``tesis_lm_index``. No purga Mesh/Curve ni toca
    materiales, para mantener Undo/Redo estable.
    """
    _unregister_lm_objs(idx)
    prefixes = (f"{idx:02d}_", f"{idx + 1}.")
    objs_to_remove = []
    for obj in list(bpy.data.objects):
        try:
            by_name = any(obj.name.startswith(prefix) for prefix in prefixes)
            by_meta = int(obj.get("tesis_lm_index", -1)) == int(idx)
        except Exception:
            by_name = any(getattr(obj, "name", "").startswith(prefix) for prefix in prefixes)
            by_meta = False
        if by_name or by_meta:
            objs_to_remove.append(obj)

    for obj in objs_to_remove:
        _safe_remove_object(obj)


def _create_or_get_landmark_material(mat_name: str, color_rgba):
    """Crear o reutilizar material con color específico con emisión para mejor visibilidad."""
    mat = bpy.data.materials.get(mat_name)
    
    # Si ya existe y está configurado, solo actualizar color
    if mat is not None and mat.use_nodes:
        mat.diffuse_color = color_rgba
        # Actualizar Base Color en bsdf
        for node in mat.node_tree.nodes:
            if node.type == "BSDF_PRINCIPLED" and "Base Color" in node.inputs:
                node.inputs["Base Color"].default_value = color_rgba
            if node.type == "EMISSION" and "Color" in node.inputs:
                node.inputs["Color"].default_value = color_rgba
        return mat
    
    mat = bpy.data.materials.new(mat_name) if mat is None else mat
    mat.use_nodes = True
    
    # Color para viewport (modo Solid) — diffuse_color es lo que muestra Solid mode
    mat.diffuse_color = color_rgba
    mat.metallic = 0.6
    mat.roughness = 0.2
    # Backface culling off para esferas pequeñas
    mat.use_backface_culling = False

    nt = mat.node_tree
    nodes = nt.nodes
    links = nt.links
    nodes.clear()

    out = nodes.new(type="ShaderNodeOutputMaterial")
    out.location = (500, 0)

    # Mix shader: 70% BSDF + 30% Emission para que brille
    mix = nodes.new(type="ShaderNodeMixShader")
    mix.location = (300, 0)
    mix.inputs[0].default_value = 0.3  # Factor: 30% Emission

    bsdf = nodes.new(type="ShaderNodeBsdfPrincipled")
    bsdf.location = (0, 100)
    if "Base Color" in bsdf.inputs:
        bsdf.inputs["Base Color"].default_value = color_rgba
    if "Roughness" in bsdf.inputs:
        bsdf.inputs["Roughness"].default_value = 0.2
    if "Metallic" in bsdf.inputs:
        bsdf.inputs["Metallic"].default_value = 0.6

    emit = nodes.new(type="ShaderNodeEmission")
    emit.location = (0, -100)
    if "Color" in emit.inputs:
        emit.inputs["Color"].default_value = color_rgba
    if "Strength" in emit.inputs:
        emit.inputs["Strength"].default_value = 0.8  # Brillo moderado

    # Conectar shaders (MixShader: inputs[1] = Shader1, inputs[2] = Shader2)
    links.new(bsdf.outputs["BSDF"], mix.inputs[1])
    links.new(emit.outputs["Emission"], mix.inputs[2])
    links.new(mix.outputs["Shader"], out.inputs["Surface"])
    
    return mat

def _get_landmark_materials():
    """Retorna materiales globales para hard (coral suave) y soft (celeste suave)."""
    mat_hard = _create_or_get_landmark_material("MAT_LM_HARD", (0.85, 0.25, 0.20, 1))
    mat_soft = _create_or_get_landmark_material("MAT_LM_SOFT", (0.20, 0.55, 0.85, 1))
    return mat_hard, mat_soft

def timestamp_now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def estado_nutricional(sexo, edad, ccu_mm):
    """Clasificación exploratoria de cribado por CCU.

    Usa los puntos de corte por sexo y edad de Fuentes Fuentes et al. (2018).
    No constituye un diagnóstico nutricional. La tabla publicada abarca 7–18 años.
    """
    puntos_ninas = {
        7: 268, 8: 271, 9: 282, 10: 285, 11: 292, 12: 312,
        13: 314, 14: 320, 15: 322, 16: 335, 17: 343, 18: 347,
    }
    puntos_ninos = {
        7: 271, 8: 280, 9: 300, 10: 302, 11: 310, 12: 320,
        13: 324, 14: 360, 15: 370, 16: 375, 17: 380, 18: 385,
    }
    ccu = _safe_float(ccu_mm, 0.0)
    edad_n = _parse_age_years(edad)
    if ccu <= 0:
        return "No aplicable"
    if ccu < 200 or ccu > 500:
        return f"Fuera de rango ({ccu:.0f} mm)"
    if edad_n is None:
        return "No evaluable (edad)"
    if edad_n not in puntos_ninas:
        return f"Sin referencia para {edad_n} años"

    sexo_s = _safe_str(sexo).strip().lower()
    if sexo_s in {"f", "femenino", "female", "mujer"}:
        punto = puntos_ninas[edad_n]
    elif sexo_s in {"m", "masculino", "male", "hombre", "varón", "varon"}:
        punto = puntos_ninos[edad_n]
    else:
        return "No evaluable (sexo)"
    return ("Cribado positivo para malnutrición por exceso"
            if ccu >= punto else "Cribado negativo para malnutrición por exceso")

def _advance_to_next_case(context):
    """Avanza al siguiente caso que cumple los filtros activos."""
    return _advance_to_next_filtered_case(context, skip_processed=False)

def _case_matches_active_filters(state: TesisState, case) -> bool:
    """Comprueba filtros de caso y, en modo ERROR, la cohorte/sesión activa."""
    error_mode = _safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() == "ERROR"

    if error_mode:
        cohort = _safe_str(getattr(state, "error_cohort", ""))
        if not bool(getattr(case, "error_study", False)):
            return False
        if cohort and _safe_str(getattr(case, "error_cohort", "")) != cohort:
            return False
        if bool(getattr(state, "error_only_pending", True)) and _error_session_complete(case, state):
            return False
    else:
        status = getattr(state, "case_status", "PENDING")
        complete = _case_is_complete(case, state=state)
        if status == "PENDING" and complete:
            return False
        if status == "PROCESSED" and not complete:
            return False

    query = _safe_str(getattr(state, "search", "")).strip().lower()
    if query:
        searchable = " ".join([
            _safe_str(case.paciente), _safe_str(case.edad),
            _safe_str(case.sexo), _safe_str(case.grupo),
            _safe_str(case.obs), _case_alert_label(case),
            _safe_str(getattr(case, "error_cohort", "")),
        ]).lower()
        if query not in searchable:
            return False

    sex_filter = getattr(state, "filter_sex", "ALL")
    sex = _safe_str(case.sexo).strip().upper()
    if sex_filter == "F" and not sex.startswith("F"):
        return False
    if sex_filter == "M" and not sex.startswith("M"):
        return False

    age = _parse_age_years(getattr(case, "edad", ""))
    age_min = int(getattr(state, "filter_age_min", 0) or 0)
    age_max = int(getattr(state, "filter_age_max", 0) or 0)
    if age_min and (age is None or age < age_min):
        return False
    if age_max and (age is None or age > age_max):
        return False

    age_band = getattr(state, "filter_age_band", "ALL")
    if age_band != "ALL":
        if age is None:
            return False
        if age_band == "6_11" and not 6 <= age <= 11:
            return False
        if age_band == "12_18" and not 12 <= age <= 18:
            return False
        if age_band == "12_14" and not 12 <= age <= 14:
            return False
        if age_band == "15_18" and not 15 <= age <= 18:
            return False

    neck_filter = getattr(state, "filter_neck", "ALL")
    if neck_filter == "YES" and not bool(case.has_neck):
        return False
    if neck_filter == "NO" and bool(case.has_neck):
        return False

    alert_filter = getattr(state, "filter_eye", "ALL")
    if alert_filter == "1" and not _case_requires_review(case):
        return False
    if alert_filter == "0" and _case_requires_review(case):
        return False

    obs_filter = getattr(state, "filter_has_obs", "ALL")
    obs_value = _safe_str(case.obs).strip()
    if obs_filter == "WITH" and (not obs_value or obs_value == "0"):
        return False
    if obs_filter == "WITHOUT" and obs_value and obs_value != "0":
        return False

    if bool(getattr(state, "filter_has_lm_excel", False)):
        if not bool(getattr(case, "summary_loaded", False)) or int(getattr(case, "summary_set", 0)) <= 0:
            return False

    return True

def _advance_to_next_filtered_case(context, skip_processed: bool = False):
    """Avanza al siguiente caso visible usando la transición atómica central."""
    state = context.scene.tesis_state
    if not state.cases:
        return False
    start = state.active_case + 1
    for idx in range(start, len(state.cases)):
        case = state.cases[idx]
        if (skip_processed
                and _safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() != "ERROR"
                and _case_is_complete(case, state=state)):
            continue
        if not _case_matches_active_filters(state, case):
            continue
        return _switch_active_case(context, idx, sync=True)
    return False

def _advance_to_prev_case(context):
    """Retrocede al caso anterior que cumple los filtros activos."""
    state = context.scene.tesis_state
    if not state.cases:
        return False
    for idx in range(state.active_case - 1, -1, -1):
        if _case_matches_active_filters(state, state.cases[idx]):
            return _switch_active_case(context, idx, sync=True)
    return False

def draw_wrapped_paragraphs(layout, text: str, context=None, width=52):
    """Dibuja párrafos independientes respetando saltos y líneas vacías."""
    panel_width = getattr(getattr(context, "region", None), "width", 0) if context else 0
    estimated = int(panel_width / 7) if panel_width else width
    wrap_width = max(28, min(52, estimated))
    for paragraph in _safe_str(text).split("\n"):
        if not paragraph.strip():
            layout.separator(factor=0.35)
            continue
        for line in textwrap.wrap(
            paragraph.strip(), width=wrap_width,
            break_long_words=False, break_on_hyphens=False,
        ):
            layout.label(text=line)


def draw_wrapped_text(layout, text: str, context=None, width=52, max_lines=4):
    """Compatibilidad: usa el envoltorio de párrafos sin truncar."""
    draw_wrapped_paragraphs(layout, text, context=context, width=width)


_VIEWPORT_OVERLAY_HANDLE = None


def _viewport_overlay_draw():
    """Dibuja una tarjeta legible del caso activo en la esquina inferior izquierda."""
    try:
        scene = bpy.context.scene
        state = getattr(scene, "tesis_state", None)
        if state is None or not state.cases or state.active_case >= len(state.cases):
            return
        case = state.cases[state.active_case]
        region = bpy.context.region
        if region is None:
            return

        patient = _safe_str(case.paciente)
        set_count = getattr(state, "lm_count_set", 0)
        missing_count = getattr(state, "lm_count_missing", 0)
        resolved_count = set_count + missing_count
        lines = [
            patient,
            f"LM {resolved_count}/{len(_OTHER_INDICES)} · {set_count} medidos · {missing_count} N/D",
        ]
        if case.has_neck:
            ccu = float(getattr(case, "ccu_mm", 0.0) or 0.0)
            lines.append(f"CCU {ccu:.1f} mm" if ccu > 0 else "CCU pendiente")
        if getattr(state, "dirty", False):
            lines.append("● SIN GUARDAR")

        font_id = 0
        blf.size(font_id, 12)
        padding = 12
        margin_x = 24
        margin_y = 40
        text_width = max(190, min(300, region.width * 0.24))
        wrapped_lines = []
        for line in lines:
            wrapped_lines.extend(textwrap.wrap(
                _safe_str(line), width=max(18, int(text_width / 7)),
                break_long_words=False, break_on_hyphens=False,
            )[:2])
        line_height = 16
        card_width = text_width + padding * 2
        card_height = len(wrapped_lines) * line_height + padding * 2
        x0 = margin_x
        y0 = margin_y
        shader = gpu.shader.from_builtin('UNIFORM_COLOR')
        vertices = ((x0, y0), (x0 + card_width, y0),
                    (x0 + card_width, y0 + card_height), (x0, y0 + card_height))
        batch = batch_for_shader(shader, 'TRI_FAN', {"pos": vertices})
        gpu.state.blend_set('ALPHA')
        shader.bind()
        shader.uniform_float("color", (0.025, 0.035, 0.05, 0.82))
        batch.draw(shader)
        gpu.state.blend_set('NONE')

        x = x0 + padding
        y = y0 + card_height - padding - line_height
        for line_number, line in enumerate(wrapped_lines):
            blf.position(font_id, x, y - line_number * line_height, 0)
            blf.color(font_id, 0.95, 0.97, 1.0, 1.0)
            blf.draw(font_id, line)
    except Exception as exc:
        log.debug("No se pudo dibujar el overlay del viewport: %s", exc, exc_info=True)


def _register_viewport_overlay():
    global _VIEWPORT_OVERLAY_HANDLE
    if _VIEWPORT_OVERLAY_HANDLE is None:
        _VIEWPORT_OVERLAY_HANDLE = bpy.types.SpaceView3D.draw_handler_add(
            _viewport_overlay_draw, (), 'WINDOW', 'POST_PIXEL'
        )


def _unregister_viewport_overlay():
    global _VIEWPORT_OVERLAY_HANDLE
    if _VIEWPORT_OVERLAY_HANDLE is not None:
        try:
            bpy.types.SpaceView3D.draw_handler_remove(_VIEWPORT_OVERLAY_HANDLE, 'WINDOW')
        except Exception:
            pass
        _VIEWPORT_OVERLAY_HANDLE = None


# =====================================================
# EXCEL IO
# =====================================================

_EXCEL_CACHE = {
    "mtime": 0.0,
    "size": 0,
    "cases": None,
    "summaries": None,
}

def _excel_file_changed(filepath: str) -> bool:
    """Retorna True si el archivo cambió desde la última lectura cacheada."""
    try:
        st = os.stat(filepath)
        changed = (st.st_mtime != _EXCEL_CACHE["mtime"]
                   or st.st_size != _EXCEL_CACHE["size"])
        if changed:
            _invalidate_excel_indexes(filepath)
        return changed
    except OSError:
        return True

def _update_cache_stamp(filepath: str):
    """Actualiza el timestamp de la caché."""
    try:
        st = os.stat(filepath)
        _EXCEL_CACHE["mtime"] = st.st_mtime
        _EXCEL_CACHE["size"] = st.st_size
    except OSError:
        pass

def invalidate_excel_cache():
    """Limpia la caché de datos Excel (llamar después de guardar al Excel)."""
    _EXCEL_CACHE["mtime"] = 0.0
    _EXCEL_CACHE["size"] = 0
    _EXCEL_CACHE["cases"] = None
    _EXCEL_CACHE["summaries"] = None


def _ensure_excel_columns(excel_path: str):
    """Asegura columnas operativas y de cohorte de error en la hoja Casos.

    La migración es aditiva: nunca elimina ni reordena columnas existentes.
    """
    required_columns = [
        COL_PROCESADO,
        COL_FECHA,
        COL_ERROR_ESTUDIO,
        COL_ERROR_COHORTE,
        COL_ERROR_ORDEN,
        COL_ERROR_ELEGIBLE,
    ]

    with open_workbook(excel_path, read_only=True, data_only=True) as wb:
        needs_rename = "Sheet1" in wb.sheetnames and SHEET_CASES not in wb.sheetnames
        sheet_name = SHEET_CASES if SHEET_CASES in wb.sheetnames else ("Sheet1" if "Sheet1" in wb.sheetnames else None)
        if sheet_name is None:
            return
        ws = wb[sheet_name]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        headers = {_safe_str(value) for value in header_row if value is not None}
        missing = [name for name in required_columns if name not in headers]

    if not missing and not needs_rename:
        return

    with open_workbook(excel_path, save=False) as wb:
        if "Sheet1" in wb.sheetnames and SHEET_CASES not in wb.sheetnames:
            wb["Sheet1"].title = SHEET_CASES
        ws = wb[SHEET_CASES]
        current = {
            _safe_str(ws.cell(1, col).value)
            for col in range(1, ws.max_column + 1)
            if ws.cell(1, col).value is not None
        }
        for name in required_columns:
            if name not in current:
                _ws_set_value(ws, 1, ws.max_column + 1, name)
                current.add(name)
        _save_workbook_safe(wb, excel_path)
        log.info("Columnas operativas/error verificadas en Excel: %s", ", ".join(required_columns))

def _read_cases_from_ws(ws, headers):
    """Lee casos válidos desde una worksheet ya abierta. Retorna lista de dicts."""
    required = [
        COL_PACIENTE, COL_OJO, COL_RUTA_BLANDO, COL_RUTA_HUESO,
        COL_INCLUIR, COL_MEJOR, COL_EDAD, COL_SEXO, COL_GRUPO, COL_PESO, COL_OBS
    ]
    for col in required:
        if col not in headers:
            raise Exception(f"Falta columna en Excel: {col}")

    i_incluir   = headers[COL_INCLUIR] - 1
    i_mejor     = headers[COL_MEJOR] - 1
    i_paciente  = headers[COL_PACIENTE] - 1
    i_ojo       = headers[COL_OJO] - 1
    i_edad      = headers[COL_EDAD] - 1
    i_sexo      = headers[COL_SEXO] - 1
    i_grupo     = headers[COL_GRUPO] - 1
    i_peso      = headers[COL_PESO] - 1
    i_obs       = headers[COL_OBS] - 1
    i_ruta_b    = headers[COL_RUTA_BLANDO] - 1
    i_ruta_h    = headers[COL_RUTA_HUESO] - 1
    i_procesado = (headers[COL_PROCESADO] - 1) if COL_PROCESADO in headers else None
    i_cuello    = (headers[COL_CUELLO] - 1) if COL_CUELLO in headers else None
    i_faltan    = (headers[COL_FALTAN] - 1) if COL_FALTAN in headers else None
    i_error_study = (headers[COL_ERROR_ESTUDIO] - 1) if COL_ERROR_ESTUDIO in headers else None
    i_error_cohort = (headers[COL_ERROR_COHORTE] - 1) if COL_ERROR_COHORTE in headers else None
    i_error_order = (headers[COL_ERROR_ORDEN] - 1) if COL_ERROR_ORDEN in headers else None
    i_error_eligible = (headers[COL_ERROR_ELEGIBLE] - 1) if COL_ERROR_ELEGIBLE in headers else None

    _TRUTHY = {"1", "1.0", "true", "sí", "si", "yes"}

    cases = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if _safe_str(row[i_incluir]).strip().lower() not in _TRUTHY:
            continue
        if _safe_str(row[i_mejor]).strip().lower() not in _TRUTHY:
            continue

        procesado_val = row[i_procesado] if i_procesado is not None else None
        has_neck = False
        if i_cuello is not None:
            has_neck = _safe_str(row[i_cuello]).strip().lower() in _TRUTHY

        faltan_lm = False
        if i_faltan is not None:
            faltan_lm = _safe_str(row[i_faltan]).strip().lower() in _TRUTHY

        cases.append({
            "paciente": _safe_str(row[i_paciente]),
            "ojo": _safe_str(row[i_ojo]),
            "edad": _safe_str(row[i_edad]),
            "sexo": _safe_str(row[i_sexo]),
            "grupo": _safe_str(row[i_grupo]),
            "peso": _safe_str(row[i_peso]),
            "obs": _safe_str(row[i_obs]),
            "ruta_blando": _safe_str(row[i_ruta_b]),
            "ruta_hueso": _safe_str(row[i_ruta_h]),
            "procesado": bool(_safe_int(procesado_val)) if procesado_val is not None else False,
            "has_neck": has_neck,
            "faltan_landmarks": faltan_lm,
            "error_study": (_safe_str(row[i_error_study]).strip().lower() in _TRUTHY) if i_error_study is not None else False,
            "error_cohort": _safe_str(row[i_error_cohort]) if i_error_cohort is not None else "",
            "error_order": _safe_int(row[i_error_order], 0) if i_error_order is not None else 0,
            "error_eligible": (_safe_str(row[i_error_eligible]).strip().lower() in _TRUTHY) if i_error_eligible is not None else False,
        })

    return cases


# =============================================================================
# CARGA DE CASOS DESDE EXCEL
# =============================================================================
# El sistema importa la lista de pacientes y sus metadatos desde la hoja
# “Casos” del archivo Excel central. En paralelo, consulta la hoja de resumen
# para recuperar el estado que ya estaba guardado del caso.
#
# En el flujo metodológico, este paso define la colección de casos del estudio y
# crea la base para la posterior importación de mallas y el marcado anatómico.
# =============================================================================
def _migrate_legacy_measurement_contexts(excel_path: str):
    """Migra solo contextos inequívocos de ``Mediciones_LM``.

    - S1 con ``ModoMedicion`` vacío -> NORMAL.
    - NORMAL fuera de S1 se CONSERVA y solo se advierte, porque puede ser una
      re-revisión histórica que requiere decisión humana.
    - Timestamp no se usa para resolver nada.
    """
    if not os.path.exists(excel_path):
        return {"migrated": 0, "normal_outside_s1": []}
    migrated = 0
    outside = []
    changed = False
    with open_workbook(excel_path, save=False) as wb:
        if SHEET_MEDICIONES not in wb.sheetnames:
            return {"migrated": 0, "normal_outside_s1": []}
        ws = wb[SHEET_MEDICIONES]
        headers = {_safe_str(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)}
        required = {"CaseID", "Sesion", "ModoMedicion", "Landmark"}
        if not required.issubset(headers):
            return {"migrated": 0, "normal_outside_s1": []}
        for row in range(2, ws.max_row + 1):
            session = _safe_str(ws.cell(row, headers["Sesion"]).value).upper()
            mode = _safe_str(ws.cell(row, headers["ModoMedicion"]).value).upper()
            if not mode and session == _PRIMARY_MEASUREMENT_SESSION:
                _ws_set_value(ws, row, headers["ModoMedicion"], "NORMAL")
                migrated += 1
                changed = True
                mode = "NORMAL"
            if mode == "NORMAL" and session != _PRIMARY_MEASUREMENT_SESSION:
                outside.append((
                    _safe_str(ws.cell(row, headers["CaseID"]).value),
                    session,
                    _safe_str(ws.cell(row, headers["Landmark"]).value),
                    row,
                ))
        if changed:
            _save_workbook_safe(wb, excel_path)
    if migrated:
        log.info("Migrados %d registros históricos S1 sin modo -> NORMAL.", migrated)
    if outside:
        cases = sorted({item[0] for item in outside if item[0]})
        log.warning(
            "%d filas NORMAL fuera de S1 en %d caso(s): %s. Se conservaron sin recodificar; Timestamp no decide.",
            len(outside), len(cases), ", ".join(cases[:12]),
        )
        print(f"[TESIS TRAZABILIDAD] {len(outside)} filas NORMAL fuera de S1: {', '.join(cases[:12])}. Revisar contexto; no se modificaron.")
    return {"migrated": migrated, "normal_outside_s1": outside}


def load_cases_and_summaries(excel_path=None):
    """Carga casos y resúmenes en una sola apertura de Excel.
    Retorna (cases, summaries)."""
    if excel_path is None:
        excel_path = get_excel_path()
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel no encontrado: {excel_path}")

    if (    _EXCEL_CACHE["cases"] is not None
        and _EXCEL_CACHE["summaries"] is not None
        and not _excel_file_changed(excel_path)):
        log.debug("Casos + resúmenes retornados desde caché.")
        return _EXCEL_CACHE["cases"], _EXCEL_CACHE["summaries"]

    _ensure_excel_columns(excel_path)
    _migrate_legacy_measurement_contexts(excel_path)

    t0 = time.perf_counter()
    n_lm = len(_OTHER_INDICES)

    with open_workbook(excel_path, read_only=True, data_only=True) as wb:
        # ── Casos ──
        if SHEET_CASES not in wb.sheetnames:
            raise Exception(f"No existe la hoja '{SHEET_CASES}' en el Excel.")
        ws = wb[SHEET_CASES]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        headers = {v: (i + 1) for i, v in enumerate(header_row) if v is not None}
        cases = _read_cases_from_ws(ws, headers)

        result = {}
        if SHEET_LANDMARKS in wb.sheetnames:
            ws_c = wb[SHEET_LANDMARKS]
            _status_idxs = []
            for idx in _OTHER_INDICES:
                base = _LM_DATA_START + idx * _LM_FIELDS_PER_LM  # 1-based
                _status_idxs.append((base - 1, base - 1 + 7))    # 0-based
            neck_status_idxs = {
                lm_id: _LM_DATA_START + idx * _LM_FIELDS_PER_LM - 1
                for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS)
                if idx in _NECK_INDEX_SET
            }

            for row in ws_c.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                pac = row[0]
                ojo = row[1]
                key = _case_pair_key(pac, ojo)
                s, o, p = 0, 0, 0
                th_vals = []
                row_len = len(row)
                for st_idx, th_idx in _status_idxs:
                    if st_idx >= row_len:
                        p += 1
                        continue
                    st = row[st_idx]
                    if st == "SET":
                        s += 1
                        if th_idx < row_len:
                            th = row[th_idx]
                            if th and float(th) > 0:
                                th_vals.append(float(th))
                    elif st == "MISSING":
                        o += 1  # no existe anatómicamente
                    else:
                        p += 1  # OMIT = pendiente
                result[key] = {
                    "set": s, "omit": o, "pend": p,
                    **_thickness_summary_from_values(th_vals),
                    "ccu_mm": 0,
                    "estado_nutricional": "",
                    "neck_ant_status": _safe_str(row[neck_status_idxs.get("neck_ant", -1)] if neck_status_idxs.get("neck_ant", -1) >= 0 and neck_status_idxs["neck_ant"] < row_len else "OMIT").upper(),
                    "neck_post_status": _safe_str(row[neck_status_idxs.get("neck_post", -1)] if neck_status_idxs.get("neck_post", -1) >= 0 and neck_status_idxs["neck_post"] < row_len else "OMIT").upper(),
                }

        if SHEET_RESULTADOS in wb.sheetnames:
            ws_r = wb[SHEET_RESULTADOS]
            hdr_row = next(ws_r.iter_rows(min_row=1, max_row=1, values_only=True))
            hdr = {v: (i + 1) for i, v in enumerate(hdr_row) if v is not None}
            if "Paciente" in hdr and "Ojo" in hdr:
                i_pac = hdr["Paciente"] - 1
                i_ojo = hdr["Ojo"] - 1
                i_ccu = (hdr["CCU_mm"] - 1) if "CCU_mm" in hdr else None
                i_est = (hdr["Estado_Nutricional"] - 1) if "Estado_Nutricional" in hdr else None
                for row in ws_r.iter_rows(min_row=2, values_only=True):
                    if not row or row[i_pac] is None:
                        continue
                    key = _case_pair_key(row[i_pac], row[i_ojo])
                    raw_ccu = row[i_ccu] if i_ccu is not None else None
                    est = row[i_est] if i_est is not None else None
                    try:
                        ccu = _safe_float(raw_ccu, default=0.0) if raw_ccu is not None else None
                    except (ValueError, TypeError):
                        ccu = None
                    if key in result:
                        if ccu:
                            result[key]["ccu_mm"] = ccu
                        if est:
                            result[key]["estado_nutricional"] = est
                    else:
                        result[key] = {
                            "set": 0, "omit": 0, "pend": 0,
                            "n_valid": 0, "avg": 0, "median": 0, "sd": 0,
                            "min": 0, "max": 0,
                            "ccu_mm": ccu if ccu else 0,
                            "estado_nutricional": _safe_str(est),
                            "neck_ant_status": "OMIT",
                            "neck_post_status": "OMIT",
                        }

        if SHEET_MEDICIONES in wb.sheetnames:
            _rebuild_error_progress_cache_from_ws(wb[SHEET_MEDICIONES])
        else:
            _ERROR_PROGRESS_CACHE.clear()

    _EXCEL_CACHE["cases"] = cases
    _EXCEL_CACHE["summaries"] = result
    _update_cache_stamp(excel_path)
    log.info("Casos + resúmenes cargados desde Excel en %.2fs (%d casos)",
             time.perf_counter() - t0, len(cases))
    return cases, result

# =====================================================
# GEOMETRÍA: BVH + closest point + raycast exterior
# =====================================================

# =============================================================================
# GEOMETRÍA Y BVH: BÚSQUEDA OPERATIVA DE SUPERFICIES
# =============================================================================
# Las mallas importadas se convierten a una representación espacial basada en
# BVH (Bounding Volume Hierarchy) para acelerar la detección de intersecciones.
#
# Este bloque es la parte automatizada del addon: la búsqueda geométrica del
# punto blando asociado a un punto óseo no depende de la intuición del usuario,
# sino de la orientación de la superficie y la trayectoria del rayo.
#
# El objetivo del cálculo es separar la decisión anatómica del observador de la
# estimación geométrica del punto de tejido blando más coherente con la malla.
# =============================================================================
def build_bvh_world(obj, depsgraph):
    """Construye un BVHTree en coordenadas WORLD a partir de un objeto evaluado."""
    obj_eval = obj.evaluated_get(depsgraph)
    me = obj_eval.to_mesh()
    mw = obj_eval.matrix_world.copy()

    verts_world = [mw @ v.co for v in me.vertices]
    polys = [p.vertices[:] for p in me.polygons]
    bvh = BVHTree.FromPolygons(verts_world, polys)

    obj_eval.to_mesh_clear()
    return bvh

def ensure_bvh_cache(context, skull_obj, soft_obj, perf=None):
    """Construye y cachea BVH Trees de skull y soft si no están cacheados o cambiaron."""
    depsgraph = context.evaluated_depsgraph_get()
    skull_sig = _bvh_object_signature(skull_obj)
    soft_sig = _bvh_object_signature(soft_obj)

    if (_BVH_CACHE["skull"] is None) or (_BVH_CACHE["skull_name"] != skull_obj.name) or (_BVH_CACHE["skull_sig"] != skull_sig):
        started = time.perf_counter()
        _BVH_CACHE["skull"] = build_bvh_world(skull_obj, depsgraph)
        if perf is not None:
            perf["construir BVH SKULL"] = time.perf_counter() - started
        _BVH_CACHE["skull_name"] = skull_obj.name
        _BVH_CACHE["skull_sig"] = skull_sig
        print(f"[TESIS BVH] SKULL REBUILD {time.perf_counter() - started:.2f} s")
    else:
        print("[TESIS BVH] SKULL CACHE HIT")

    if (_BVH_CACHE["soft"] is None) or (_BVH_CACHE["soft_name"] != soft_obj.name) or (_BVH_CACHE["soft_sig"] != soft_sig):
        started = time.perf_counter()
        _BVH_CACHE["soft"] = build_bvh_world(soft_obj, depsgraph)
        if perf is not None:
            perf["construir BVH SOFT"] = time.perf_counter() - started
        _BVH_CACHE["soft_name"] = soft_obj.name
        _BVH_CACHE["soft_sig"] = soft_sig
        print(f"[TESIS BVH] SOFT REBUILD {time.perf_counter() - started:.2f} s")
    else:
        print("[TESIS BVH] SOFT CACHE HIT")

def closest_point_on_obj_world(obj, point_world, depsgraph):
    """Proyecta un punto world al punto más cercano en la malla del objeto.
    Retorna (location_world, normal_world) o (None, None) si falla.
    """
    obj_eval = obj.evaluated_get(depsgraph)
    mw = obj_eval.matrix_world
    mw_inv = mw.inverted()
    point_local = mw_inv @ point_world
    ok, loc_local, normal_local, face = obj_eval.closest_point_on_mesh(point_local)
    if not ok:
        return None, None
    loc_world = mw @ loc_local
    normal_world = (mw.to_3x3() @ normal_local).normalized()
    return loc_world, normal_world

def obj_center_world(obj):
    """Calcula el centro aproximado del objeto en world space usando bounding box."""
    bb = obj.bound_box
    if not bb:
        return obj.matrix_world.translation.copy()
    c = Vector((0, 0, 0))
    for v in bb:
        c += obj.matrix_world @ Vector(v)
    return c / 8.0

def outward_normal_at_point(skull_obj, point_world, normal_world):
    """Asegura que la normal apunte hacia afuera del cráneo (lejos del centro)."""
    center = obj_center_world(skull_obj)
    v = (point_world - center)
    if v.length == 0:
        return normal_world.normalized()
    # Si normal apunta hacia el centro, flip
    if normal_world.dot(v) < 0:
        return (-normal_world).normalized()
    return normal_world.normalized()


def ray_cast_last_hit(bvh, origin, direction, max_dist=500.0, max_steps=64):
    """Lanza un rayo y retorna el úLTIMO punto de intersección con la BVH.
    Útil para encontrar la salida opuesta del rayo a través de la malla.
    Avanza en pasos pequeños (0.01mm) tras cada hit hasta no encontrar más.
    """
    dirn = direction.normalized()
    last_hit = None
    o = origin.copy()
    remaining = max_dist

    for _ in range(max_steps):
        hit = bvh.ray_cast(o, dirn, remaining)
        if hit[0] is None:
            break

        p_hit = hit[0]
        last_hit = p_hit

        step = (p_hit - o).length
        o = p_hit + dirn * 0.01
        remaining -= step
        if remaining <= 0:
            break

    return last_hit


def ray_cast_all_hits(bvh, origin, direction, max_dist=200.0, epsilon=0.05, max_hits=32):
    """Retorna todas las intersecciones de un rayo, ordenadas desde origin."""
    if bvh is None or direction.length <= 1e-9 or max_hits <= 0 or max_dist <= 0:
        return []

    dirn = direction.normalized()
    origin_world = Vector(origin)
    current_origin = origin_world.copy()
    remaining = float(max_dist)
    advance = max(float(epsilon), 1e-6)
    hits = []

    for _ in range(int(max_hits)):
        hit_point, hit_normal, face_index, _hit_distance = bvh.ray_cast(
            current_origin, dirn, remaining
        )
        if hit_point is None:
            break

        point = Vector(hit_point)
        distance = (point - origin_world).length
        hits.append({
            "point": point.copy(),
            "distance": float(distance),
            "normal": hit_normal.copy() if hit_normal is not None else None,
            "face_index": int(face_index) if face_index is not None and face_index >= 0 else None,
        })

        travelled = (point - current_origin).length
        step = max(advance, travelled + advance)
        remaining -= step
        if remaining <= 0:
            break
        current_origin = point + dirn * advance

    return hits


def is_soft_candidate_visible_from_view(bvh_soft, view_origin, candidate, tolerance_mm=0.5):
    """Comprueba que el candidato sea el primer impacto visto desde el viewport."""
    direction = Vector(candidate) - Vector(view_origin)
    distance = direction.length
    if distance <= 1e-9:
        return False
    hits = ray_cast_all_hits(
        bvh_soft, view_origin, direction.normalized(),
        max_dist=distance + max(float(tolerance_mm), 0.01),
        epsilon=0.05,
        max_hits=1,
    )
    return bool(hits) and (hits[0]["point"] - Vector(candidate)).length <= tolerance_mm

def _get_neck_material():
    """Retorna material verde para landmarks de cuello."""
    return _create_or_get_landmark_material("MAT_LM_NECK", (0.15, 0.85, 0.30, 1))


_LM_SPHERE_MESH_NAMES = {
    "HARD": "TESIS_LM_SPHERE_HARD",
    "SOFT": "TESIS_LM_SPHERE_SOFT",
    "NECK": "TESIS_LM_SPHERE_NECK",
    "DIAGNOSTIC": "TESIS_LM_SPHERE_DIAGNOSTIC",
}


def _get_shared_landmark_mesh(role):
    mesh_name = _LM_SPHERE_MESH_NAMES[role]
    mesh = bpy.data.meshes.get(mesh_name)
    if mesh is not None:
        return mesh
    bm = bmesh.new()
    try:
        bmesh.ops.create_uvsphere(bm, u_segments=10, v_segments=6, radius=1.0)
        mesh = bpy.data.meshes.new(mesh_name)
        bm.to_mesh(mesh)
        mesh.update()
        mesh["tesis_owned"] = True
        mesh["tesis_role"] = f"LM_{role}"
        return mesh
    finally:
        bm.free()


def _remove_soft_surface_constraint(soft_obj) -> bool:
    """Elimina solo el constraint legacy creado por el addon."""
    removed = False
    if soft_obj is None:
        return removed
    for constraint in list(soft_obj.constraints):
        if constraint.name == "TESIS_SOFT_SURFACE":
            soft_obj.constraints.remove(constraint)
            removed = True
    return removed


def create_landmark_sphere(name: str, location, radius_mm: float, color_rgba, collection_name: str):
    """
    Crea esfera pequeña vía bmesh (sin bpy.ops).
    Radio 0.8 mm, pocas caras para aspecto limpio.
    """
    r, g, b = color_rgba[0], color_rgba[1], color_rgba[2]
    role = "NECK" if g > r and g > b else ("HARD" if r > 0.5 else "SOFT")
    if abs(r - g) < 1e-6 and abs(g - b) < 1e-6:
        role = "DIAGNOSTIC"
    obj = bpy.data.objects.new(name, _get_shared_landmark_mesh(role))
    obj.location = Vector(location) if not isinstance(location, Vector) else location
    obj.scale = (radius_mm, radius_mm, radius_mm)

    # Mostrar nombres según preferencia del usuario
    state = getattr(getattr(bpy.context, "scene", None), "tesis_state", None)
    obj.show_name = False
    obj.show_in_front = False

    # Seleccionar material por tipo de landmark
    if g > r and g > b:  # Verde = cuello
        mat = _get_neck_material()
    elif r > 0.5:  # Rojo = hard
        mat = _get_landmark_materials()[0]
    else:  # Azul = soft
        mat = _get_landmark_materials()[1]

    if not obj.data.materials:
        obj.data.materials.append(mat)

    # Color del objeto para viewport Solid > Color = Object
    obj.color = color_rgba

    # Enviar a colección correcta
    col = bpy.data.collections.get(collection_name)
    if col is None:
        col = get_or_create_collection(collection_name)

    move_object_to_collection(obj, col)

    return obj

_MALLAS_CACHE = {"key": None, "time": 0.0, "result": (None, None)}
_MALLAS_CACHE_TTL = 0.2  # seg. — evita re-escanear bpy.data.objects en cada redraw de panel
_DEFERRED_CLEANUP_PENDING = False
_DEFERRED_CLEANUP_TIMER = False


def invalidate_mallas_cache():
    """Descarta referencias Object antes de desvincular objetos del caso."""
    _MALLAS_CACHE["key"] = None
    _MALLAS_CACHE["time"] = 0.0
    _MALLAS_CACHE["result"] = (None, None)

def obtener_mallas_paciente(*, quiet: bool = False):
    """Retorna las mallas SKULL y SOFT del caso activo.

    Prioriza las propiedades ``tesis_case_key``/``tesis_role`` y usa el nombre
    como respaldo. El resultado se conserva brevemente para evitar recorrer la
    colección cada vez que Blender redibuja el panel.
    """
    col = bpy.data.collections.get("TESIS_MESHES")
    if col is None:
        return None, None

    state = getattr(getattr(bpy.context, "scene", None), "tesis_state", None)
    if not state or not getattr(state, "cases", None):
        if not quiet:
            log.warning("obtener_mallas_paciente: no hay caso activo.")
        return None, None

    try:
        case = state.cases[state.active_case]
        case_key = _current_case_key(case)
        expected_skull = f"SKULL_{case.paciente}_{case.ojo}"
        expected_soft = f"SOFT_{case.paciente}_{case.ojo}"
    except (IndexError, AttributeError):
        if not quiet:
            log.warning("obtener_mallas_paciente: no se pudo leer el caso activo.")
        return None, None

    now = time.monotonic()
    if (_MALLAS_CACHE["key"] == case_key
            and now - _MALLAS_CACHE["time"] < _MALLAS_CACHE_TTL):
        cached_skull, cached_soft = _MALLAS_CACHE["result"]
        try:
            skull_ok = cached_skull is None or bpy.data.objects.get(cached_skull.name) is cached_skull
            soft_ok = cached_soft is None or bpy.data.objects.get(cached_soft.name) is cached_soft
            if skull_ok and soft_ok:
                return cached_skull, cached_soft
        except ReferenceError:
            invalidate_mallas_cache()

    skull = None
    soft = None
    for obj in col.objects:
        try:
            if _safe_str(obj.get("tesis_case_key")) != case_key:
                continue
            role = _safe_str(obj.get("tesis_role")).upper()
            if role == "SKULL":
                skull = obj
            elif role == "SOFT":
                soft = obj
        except Exception:
            continue

    if skull is None:
        skull = col.objects.get(expected_skull) or bpy.data.objects.get(expected_skull)
    if soft is None:
        soft = col.objects.get(expected_soft) or bpy.data.objects.get(expected_soft)

    result = (skull, soft)
    _MALLAS_CACHE.update({"key": case_key, "time": now, "result": result})

    if (skull is None or soft is None) and not quiet:
        log.warning(
            "No se encontraron ambas mallas del caso activo (esperado: %s, %s).",
            expected_skull, expected_soft,
        )
    return result


def _remove_curve_point_driver(point, axis: int):
    try:
        point.driver_remove("co", axis)
        return
    except Exception:
        pass
    try:
        path = point.path_from_id("co")
        point.id_data.driver_remove(path, axis)
    except Exception:
        pass


def _add_world_location_driver(point, axis: int, target_obj) -> bool:
    """Vincula una coordenada de un punto de curva a la posición WORLD de un objeto."""
    if point is None or target_obj is None:
        return False
    _remove_curve_point_driver(point, axis)
    try:
        try:
            fcurve = point.driver_add("co", axis)
        except Exception:
            path = point.path_from_id("co")
            fcurve = point.id_data.driver_add(path, axis)
        driver = fcurve.driver
        # AVERAGE con una sola variable evita expresiones Python y restricciones
        # de seguridad de Auto Run Scripts.
        driver.type = 'AVERAGE'
        while driver.variables:
            driver.variables.remove(driver.variables[0])
        var = driver.variables.new()
        var.name = "loc"
        var.type = 'TRANSFORMS'
        target = var.targets[0]
        target.id = target_obj
        target.transform_type = ('LOC_X', 'LOC_Y', 'LOC_Z')[axis]
        target.transform_space = 'WORLD_SPACE'
        return True
    except Exception as exc:
        log.debug("No se pudo crear driver de línea para %s: %s", getattr(target_obj, 'name', '?'), exc)
        return False


def _bind_line_to_landmark_objects(line_obj, hard_obj, soft_obj) -> bool:
    """Hace que ambos extremos de la línea sigan HARD y SOFT nativamente.

    Los drivers se evalúan dentro de Blender, por lo que la línea acompaña al SOFT
    durante ``G`` sin handlers de depsgraph ni operadores modales superpuestos.
    """
    if line_obj is None or line_obj.type != 'CURVE' or hard_obj is None or soft_obj is None:
        return False
    if not line_obj.data.splines or len(line_obj.data.splines[0].points) < 2:
        return False

    point_hard = line_obj.data.splines[0].points[0]
    point_soft = line_obj.data.splines[0].points[1]
    ok = True
    for axis in range(3):
        ok = _add_world_location_driver(point_hard, axis, hard_obj) and ok
        ok = _add_world_location_driver(point_soft, axis, soft_obj) and ok

    try:
        line_obj["tesis_hard_name"] = hard_obj.name
        line_obj["tesis_soft_name"] = soft_obj.name
        line_obj["tesis_line_bound"] = bool(ok)
    except Exception:
        pass
    return ok


def _evaluated_world_location(obj, depsgraph):
    """Obtiene la posición WORLD final, incluyendo constraints y parenting."""
    if obj is None:
        return None
    try:
        return obj.evaluated_get(depsgraph).matrix_world.translation.copy()
    except Exception:
        try:
            return obj.matrix_world.translation.copy()
        except Exception:
            return None


def create_line_object(name, p0, p1, hard_obj=None, soft_obj=None):
    """Crea una línea HARD–SOFT y, si puede, vincula sus extremos a los landmarks."""
    curve_data = bpy.data.curves.new(name=name, type='CURVE')
    curve_data["tesis_owned"] = True
    curve_data["tesis_role"] = "LINE"
    curve_data.dimensions = '3D'
    spline = curve_data.splines.new(type='POLY')
    spline.points.add(1)
    spline.points[0].co = (p0.x, p0.y, p0.z, 1.0)
    spline.points[1].co = (p1.x, p1.y, p1.z, 1.0)

    obj = bpy.data.objects.new(name, curve_data)
    obj.show_in_front = True
    state = getattr(getattr(bpy.context, "scene", None), "tesis_state", None)
    obj.show_name = False
    obj.hide_select = True
    obj.location = (0.0, 0.0, 0.0)
    obj.rotation_euler = (0.0, 0.0, 0.0)
    obj.scale = (1.0, 1.0, 1.0)

    cols = ensure_tesis_collections()
    move_object_to_collection(obj, cols["lines"])

    line_bound = False
    if hard_obj is not None and soft_obj is not None:
        line_bound = _bind_line_to_landmark_objects(obj, hard_obj, soft_obj)

    # Los extremos se actualizan explícitamente al crear o mover el landmark.
    try:
        obj["tesis_hard_name"] = hard_obj.name if hard_obj else ""
        obj["tesis_soft_name"] = soft_obj.name if soft_obj else ""
        obj["tesis_line_bound"] = bool(line_bound)
    except Exception:
        pass
    return obj


def update_line_endpoints(line_obj, p0, p1):
    """Fallback manual para líneas sin drivers o para reparación/guardado."""
    if line_obj is None or line_obj.type != 'CURVE':
        return False
    curve_data = line_obj.data
    if not curve_data.splines or len(curve_data.splines[0].points) < 2:
        return False
    try:
        inv = line_obj.matrix_world.inverted_safe()
        p0_local = inv @ Vector(p0)
        p1_local = inv @ Vector(p1)
        spline = curve_data.splines[0]
        spline.points[0].co = (p0_local.x, p0_local.y, p0_local.z, 1.0)
        spline.points[1].co = (p1_local.x, p1_local.y, p1_local.z, 1.0)
        curve_data.update_tag()
        line_obj.update_tag(refresh={'DATA'})
        return True
    except Exception as exc:
        log.debug("No se pudieron actualizar extremos de %s: %s", getattr(line_obj, 'name', '?'), exc)
        return False


def _build_coordenadas_header():
    """Header para hoja Coordenadas: posiciones 3D crudas por landmark."""
    cols = ["Paciente", "Ojo"]
    for lm_id, _soft_id, _label in LANDMARKS_ESPERADOS:
        cols.extend([
            f"{lm_id}_St",
            f"{lm_id}_HX", f"{lm_id}_HY", f"{lm_id}_HZ",
            f"{lm_id}_SX", f"{lm_id}_SY", f"{lm_id}_SZ",
            f"{lm_id}_Th",
        ])
    cols.extend(["UnitScale", "Timestamp"])
    return cols


def _apply_header_style(ws, num_cols):
    """Aplica negrita + fondo gris claro a la primera fila del sheet."""
    fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    bold = Font(bold=True)
    for col in range(1, num_cols + 1):
        cell = ws.cell(1, col)
        cell.font = bold
        cell.fill = fill


def _build_resultados_header():
    """Header para hoja Resultados: datos demográficos + grosor (mm) por landmark + CCU + estado."""
    cols = ["Paciente", "Ojo", "Edad", "Sexo"]
    for lm_id, _soft_id, _label in (LANDMARKS_ESPERADOS[idx] for idx in _OTHER_INDICES):
        cols.append(f"{lm_id}_mm")
    cols.extend([
        "CCU_mm", "Estado_Nutricional",
        "Timestamp",
    ])
    return cols


_LM_HEADER = _build_coordenadas_header()
_RES_HEADER = _build_resultados_header()
_LM_FIELDS_PER_LM = 8  # St, HX, HY, HZ, SX, SY, SZ, Th
_LM_DATA_START = 3      # columna C (1-indexed) = después de Paciente, Ojo
_RES_DATA_START = 5     # columna E para Resultados (después de Paciente, Ojo, Edad, Sexo)


def ensure_landmarks_sheet(wb):
    """Crear o migrar hoja Coordenadas (antes 'Landmarks') al formato horizontal."""
    # Migrar nombre antiguo "Sheet1" → "Casos"
    if "Sheet1" in wb.sheetnames and SHEET_CASES not in wb.sheetnames:
        wb["Sheet1"].title = SHEET_CASES

    # Migrar nombre antiguo "Landmarks" → "Coordenadas"
    if "Landmarks" in wb.sheetnames and SHEET_LANDMARKS not in wb.sheetnames:
        ws_old = wb["Landmarks"]
        ws_old.title = SHEET_LANDMARKS

    if SHEET_LANDMARKS not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_LANDMARKS)
        for i, h in enumerate(_LM_HEADER, 1):
            ws.cell(1, i).value = h
        _apply_header_style(ws, len(_LM_HEADER))
        return ws

    ws = wb[SHEET_LANDMARKS]
    first_cell = ws.cell(1, 3).value or ""
    # Si la columna C es "Landmark", es formato vertical antiguo → borrar y recrear
    if first_cell == "Landmark":
        ws.delete_rows(1, ws.max_row)
        for i, h in enumerate(_LM_HEADER, 1):
            ws.cell(1, i).value = h
        _apply_header_style(ws, len(_LM_HEADER))
        return ws

    # Validar / actualizar header
    current = [ws.cell(1, c).value for c in range(1, len(_LM_HEADER) + 1)]
    if current != _LM_HEADER:
        for i, h in enumerate(_LM_HEADER, 1):
            ws.cell(1, i).value = h
    _apply_header_style(ws, len(_LM_HEADER))
    return ws


def ensure_resultados_sheet(wb):
    """Crear o migrar hoja Resultados al formato con grosor por landmark."""
    if SHEET_RESULTADOS not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_RESULTADOS)
        for i, h in enumerate(_RES_HEADER, 1):
            ws.cell(1, i).value = h
        _apply_header_style(ws, len(_RES_HEADER))
        return ws

    ws = wb[SHEET_RESULTADOS]
    # Detectar formatos antiguos:
    #  - col 3 == "SET" → formato resumen viejo
    #  - col 3 == "g_mm" → formato sin Edad/Sexo
    old_c3 = _safe_str(ws.cell(1, 3).value)
    if old_c3 in ("SET", "g_mm"):
        # Formato antiguo → borrar datos y recrear con nuevo header
        ws.delete_rows(1, ws.max_row)
        for i, h in enumerate(_RES_HEADER, 1):
            ws.cell(1, i).value = h
        _apply_header_style(ws, len(_RES_HEADER))
        return ws

    # Validar / actualizar header
    current = [ws.cell(1, c).value for c in range(1, len(_RES_HEADER) + 1)]
    if current != _RES_HEADER:
        for i, h in enumerate(_RES_HEADER, 1):
            ws.cell(1, i).value = h
    _apply_header_style(ws, len(_RES_HEADER))
    return ws


def ensure_disponibilidad_sheet(wb):
    """Crea o actualiza la hoja de disponibilidad de landmarks para resumen visible."""
    if SHEET_DISPONIBILIDAD not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_DISPONIBILIDAD)
        headers = [
            "Paciente", "Edad", "Sexo", "Grupo", "Procesado", "TieneCuello", "Observaciones"
        ]
        for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS[:26], 1):
            headers.append(lm_id)
        headers.extend(["N_Medidos", "N_NoDisponibles", "N_Pendientes", "FacialCompleto", "Cobertura_Medida_pct", "Resolucion_pct", "EstadoCaso", "Neck_ant_estado", "Neck_post_estado", "CCU_mm", "UltimaActualizacion"])
        for i, h in enumerate(headers, 1):
            ws.cell(1, i).value = h
        _apply_header_style(ws, len(headers))
        return ws

    ws = wb[SHEET_DISPONIBILIDAD]
    expected = [
        "Paciente", "Edad", "Sexo", "Grupo", "Procesado", "TieneCuello", "Observaciones"
    ]
    for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS[:26], 1):
        expected.append(lm_id)
    expected.extend(["N_Medidos", "N_NoDisponibles", "N_Pendientes", "FacialCompleto", "Cobertura_Medida_pct", "Resolucion_pct", "EstadoCaso", "Neck_ant_estado", "Neck_post_estado", "CCU_mm", "UltimaActualizacion"])
    current = [ws.cell(1, c).value for c in range(1, len(expected) + 1)]
    if current != expected:
        for i, h in enumerate(expected, 1):
            ws.cell(1, i).value = h
    _apply_header_style(ws, len(expected))
    return ws


def _build_mediciones_header():
    return [
        "CaseID",
        "Paciente",
        "AlertaCaso",
        "Edad",
        "Sexo",
        "Grupo",
        "Peso",
        "TieneCuello",
        "FaltanLandmarks",
        "Observaciones",
        "Observador",
        "Sesion",
        "Replica",
        "Timestamp",  # bitácora; no identifica la medición ni resuelve duplicados
        "Landmark",
        "GrupoLM",
        "Lado",
        "ParID",
        "Status",
        "HX", "HY", "HZ",
        "SX", "SY", "SZ",
        "Thickness_mm",
        "DeltaBilateral_mm",
        "ModoMedicion",
        "CohorteError",
        "CCU_mm",
    ]


_MED_HEADER = _build_mediciones_header()


def ensure_mediciones_sheet(wb):
    if SHEET_MEDICIONES not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_MEDICIONES)
        for i, h in enumerate(_MED_HEADER, 1):
            _ws_set_value(ws, 1, i, h)
        _apply_header_style(ws, len(_MED_HEADER))
        return ws

    ws = wb[SHEET_MEDICIONES]
    current = [ws.cell(1, c).value for c in range(1, len(_MED_HEADER) + 1)]
    if current != _MED_HEADER:
        for i, h in enumerate(_MED_HEADER, 1):
            _ws_set_value(ws, 1, i, h)
    _apply_header_style(ws, len(_MED_HEADER))
    return ws


def _lm_group_side_pair(lm_id: str):
    lid = _safe_str(lm_id)
    lid_norm = _normalize_lm(lid)

    if lid_norm in _NECK_IDS:
        return "neck", "MID", lid

    if lid.endswith("-izq"):
        return "bilateral", "L", lid[:-4]

    if lid.endswith("-der"):
        return "bilateral", "R", lid[:-4]

    return "central", "MID", lid


def _append_mediciones_long(wb, case_data, lm_data, excel_path=None):
    ws = ensure_mediciones_sheet(wb)
    # Timestamp se conserva solo para auditoría humana; la clave lógica está abajo.
    ts = timestamp_now()
    case_id = _safe_str(case_data.get("case_id") or case_data.get("paciente"))
    observer = _safe_str(case_data.get("observer_id", "OBS1")).upper()
    session = _safe_str(case_data.get("session_id", "S1")).upper()
    replica = int(case_data.get("replicate_id", 1))
    mode = _safe_str(case_data.get("measurement_mode", "NORMAL")).upper() or "NORMAL"
    cohort = _safe_str(case_data.get("error_cohort", "")) if mode == "ERROR" else ""
    ccu_mm = _safe_float(case_data.get("ccu_mm", 0.0), 0.0)

    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    def key_builder(r):
        row_mode = _safe_str(ws.cell(r, headers.get("ModoMedicion", 0)).value).upper() if headers.get("ModoMedicion") else "NORMAL"
        if not row_mode:
            row_mode = "NORMAL"
        row_cohort = _safe_str(ws.cell(r, headers.get("CohorteError", 0)).value) if headers.get("CohorteError") else ""
        return (
            _safe_str(ws.cell(r, headers["CaseID"]).value),
            _safe_str(ws.cell(r, headers["Observador"]).value).upper(),
            _safe_str(ws.cell(r, headers["Sesion"]).value).upper(),
            _safe_int(ws.cell(r, headers["Replica"]).value, 1),
            _safe_str(ws.cell(r, headers["Landmark"]).value),
            row_mode,
            row_cohort if row_mode == "ERROR" else "",
        )
    existing = (_cached_sheet_index(excel_path, ws, key_builder)
                if excel_path else {
                    key_builder(r): r for r in range(2, ws.max_row + 1)
                })

    for idx, lm in enumerate(lm_data):
        lm_id = LANDMARKS_ESPERADOS[idx][0]
        grupo_lm, lado, par_id = _lm_group_side_pair(lm_id)
        status = _safe_str(lm.get("status", "OMIT")).upper()
        valid = status == "SET"
        is_neck = _normalize_lm(lm_id) in _NECK_IDS

        delta_bilat = None
        pair_idx = _BILATERAL_PAIR_MAP.get(idx)
        if pair_idx is not None and pair_idx < len(lm_data):
            pair_lm = lm_data[pair_idx]
            if valid and pair_lm.get("status") == "SET":
                delta_bilat = round(abs(float(lm.get("thickness", 0.0)) - float(pair_lm.get("thickness", 0.0))), 4)

        key = (case_id, observer, session, replica, lm_id, mode, cohort if mode == "ERROR" else "")
        existing_row = existing.get(key)
        row = existing_row if isinstance(existing_row, int) else ws.max_row + 1
        values = [
            case_id, case_data["paciente"], case_data.get("alerta_caso", 0),
            case_data.get("edad", ""), case_data.get("sexo", ""), case_data.get("grupo", ""),
            case_data.get("peso", ""), 1 if case_data.get("has_neck", False) else 0,
            1 if case_data.get("faltan_landmarks", False) else 0, case_data.get("obs", ""),
            observer, session, replica, ts, lm_id, grupo_lm, lado, par_id, status,
            lm.get("hx") if valid and not is_neck else None,
            lm.get("hy") if valid and not is_neck else None,
            lm.get("hz") if valid and not is_neck else None,
            lm.get("x") if valid else None, lm.get("y") if valid else None, lm.get("z") if valid else None,
            lm.get("thickness") if valid and not is_neck else None,
            delta_bilat,
            mode,
            cohort,
            ccu_mm if ccu_mm > 0 else None,
        ]
        for col, val in enumerate(values, 1):
            _ws_set_value(ws, row, col, val)
        existing[key] = row


def _build_patient_row_index(ws, col_pac=1, col_ojo=2):
    """Construye un dict {(paciente, ojo): row_num} para búsqueda O(1).
    Útil cuando se escribe al Excel frecuentemente o con muchos pacientes.

    Args:
        col_pac: Columna 1-based del campo Paciente (default 1 para hojas creadas por el addon).
        col_ojo: Columna 1-based del campo Ojo (default 2).
    """
    idx = {}
    for r in range(2, ws.max_row + 1):
        p = _safe_str(ws.cell(r, col_pac).value)
        o = _safe_str(ws.cell(r, col_ojo).value)
        if p:  # solo filas con datos
            idx[(p, o)] = r
    return idx


def _build_case_row_index(ws, headers):
    """Agrupa filas Casos por paciente/ojo para conservar la prioridad operativa."""
    index = {}
    col_pac = headers.get(COL_PACIENTE, 1)
    col_ojo = headers.get(COL_OJO, 2)
    for row in range(2, ws.max_row + 1):
        key = (_safe_str(ws.cell(row, col_pac).value),
               _safe_str(ws.cell(row, col_ojo).value))
        if key[0]:
            index.setdefault(key, []).append(row)
    return index


def _find_case_row_prefer_operational(ws, paciente, ojo, headers, row_index=None):
    """Busca la fila de Casos correcta cuando paciente/ojo está duplicado.

    Prioriza una fila con Incluir=1 y Mejor=1. Esto evita que una serie
    alternativa (Mejor=0) reciba Procesado/Observaciones por accidente.
    """
    col_pac = headers.get(COL_PACIENTE, 1)
    col_ojo = headers.get(COL_OJO, 2)
    col_incluir = headers.get(COL_INCLUIR)
    col_mejor = headers.get(COL_MEJOR)
    target_ojo = _safe_str(ojo)
    matches = []
    truthy_values = {"1", "1.0", "true", "sí", "si", "yes"}
    candidate_rows = (row_index.get((_safe_str(paciente), target_ojo), [])
                      if row_index is not None else range(2, ws.max_row + 1))
    for r in candidate_rows:
        p = _safe_str(ws.cell(r, col_pac).value)
        o = _safe_str(ws.cell(r, col_ojo).value)
        if p != _safe_str(paciente) or o != target_ojo:
            continue
        incluir = (_safe_str(ws.cell(r, col_incluir).value).lower() in truthy_values) if col_incluir else False
        mejor = (_safe_str(ws.cell(r, col_mejor).value).lower() in truthy_values) if col_mejor else False
        score = (2 if mejor else 0) + (1 if incluir else 0)
        matches.append((score, r))
    if not matches:
        return None
    matches.sort(key=lambda item: (item[0], -item[1]), reverse=True)
    return matches[0][1]


def _find_patient_row(ws, paciente, ojo, _index=None):
    """Busca fila existente para paciente+ojo. Retorna nro de fila o None.
    
    Args:
        _index: dict pre-construido con _build_patient_row_index para evitar
                escaneo lineal repetido. Si None, hace búsqueda lineal.
    """
    if _index is not None:
        return _index.get(_case_pair_key(paciente, ojo))
    for r in range(2, ws.max_row + 1):
        p = _safe_str(ws.cell(r, 1).value)
        o = _safe_str(ws.cell(r, 2).value)
        if _case_pair_key(p, o) == _case_pair_key(paciente, ojo):
            return r
    return None


def _serialize_landmarks(landmarks):
    """Serializa landmarks de bpy a lista de dicts puros (thread-safe)."""
    data = []
    for lm in landmarks:
        status = _safe_str(lm.status).upper()
        is_neck = _normalize_lm(getattr(lm, "name", "")) in _NECK_IDS
        valid = status == "SET"
        data.append({
            "status": status,
            "hx": round(float(getattr(lm, "hx", 0.0)), 4) if valid and not is_neck else None,
            "hy": round(float(getattr(lm, "hy", 0.0)), 4) if valid and not is_neck else None,
            "hz": round(float(getattr(lm, "hz", 0.0)), 4) if valid and not is_neck else None,
            "x": round(float(lm.x), 4) if valid else None,
            "y": round(float(lm.y), 4) if valid else None,
            "z": round(float(lm.z), 4) if valid else None,
            "thickness": round(float(getattr(lm, "thickness", 0.0)), 4) if valid and not is_neck else None,
        })
    return data


def _serialize_case(case, state=None):
    """Serializa datos del caso de bpy a dict puro (thread-safe)."""
    data = {
        "paciente": case.paciente,
        "case_id": case.paciente,
        "ojo": case.ojo,  # legacy, solo para compatibilidad con hojas antiguas
        "alerta_caso": _case_alert_int(case),
        "edad": case.edad,
        "sexo": case.sexo,
        "grupo": case.grupo,
        "peso": case.peso,
        "obs": case.obs,
        "has_neck": bool(getattr(case, "has_neck", False)),
        "faltan_landmarks": bool(getattr(case, "faltan_landmarks", False)),
        "ccu_mm": float(getattr(case, "ccu_mm", 0.0)),
        "estado_nutricional": estado_nutricional(case.sexo, case.edad, getattr(case, "ccu_mm", 0.0)) if bool(getattr(case, "has_neck", False)) else "No aplicable",
        "error_study": bool(getattr(case, "error_study", False)),
        "error_case_cohort": _safe_str(getattr(case, "error_cohort", "")),
        "error_order": int(getattr(case, "error_order", 0) or 0),
    }

    if state is not None:
        data["observer_id"] = _safe_str(getattr(state, "observer_id", "OBS1")).upper() or "OBS1"
        data["session_id"] = _safe_str(getattr(state, "session_id", "S1")).upper() or "S1"
        data["replicate_id"] = int(getattr(state, "replicate_id", 1) or 1)
        data["measurement_mode"] = _safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() or "NORMAL"
        data["error_cohort"] = (
            _safe_str(getattr(state, "error_cohort", ""))
            if data["measurement_mode"] == "ERROR" else ""
        )
    else:
        data["observer_id"] = "OBS1"
        data["session_id"] = "S1"
        data["replicate_id"] = 1
        data["measurement_mode"] = "NORMAL"
        data["error_cohort"] = ""

    return data


def _write_disponibilidad_lm_sheet(wb, case_data, lm_data):
    """Actualiza la hoja de disponibilidad de landmarks con el resumen visible del caso."""
    ws = ensure_disponibilidad_sheet(wb)
    summary = ensure_resumen_lm_sheet(wb)
    if (max(1, summary.max_row) < 27
            or all(_safe_int(summary.cell(row, 5).value, 0) == 0 for row in range(2, 28))):
        _write_resumen_lm_sheet(wb)
    paciente = case_data["paciente"]
    old_values = []
    row_index = getattr(wb, "_tesis_disponibilidad_index", None)
    if row_index is None:
        row_index = {}
        for candidate in range(2, ws.max_row + 1):
            key = _safe_str(ws.cell(candidate, 1).value)
            if key:
                row_index.setdefault(key, candidate)
        setattr(wb, "_tesis_disponibilidad_index", row_index)
    row = row_index.get(_safe_str(paciente))
    if row is None:
        row = next((candidate for candidate in range(2, ws.max_row + 1)
                    if _safe_str(ws.cell(candidate, 1).value) == _safe_str(paciente)), None)
    if row is not None:
        old_values = [_safe_str(ws.cell(row, 8 + pos).value) for pos in range(26)]
    if row is None:
        row = ws.max_row + 1
        ws.cell(row, 1).value = paciente
        ws.cell(row, 2).value = case_data.get("edad", "")
        ws.cell(row, 3).value = case_data.get("sexo", "")
        ws.cell(row, 4).value = case_data.get("grupo", "")
        ws.cell(row, 5).value = 1 if bool(case_data.get("procesado", False)) else 0
        ws.cell(row, 6).value = 1 if bool(case_data.get("has_neck", False)) else 0
        ws.cell(row, 7).value = case_data.get("obs", "")
    else:
        ws.cell(row, 1).value = paciente
        ws.cell(row, 2).value = case_data.get("edad", "")
        ws.cell(row, 3).value = case_data.get("sexo", "")
        ws.cell(row, 4).value = case_data.get("grupo", "")
        ws.cell(row, 5).value = 1 if bool(case_data.get("procesado", False)) else 0
        ws.cell(row, 6).value = 1 if bool(case_data.get("has_neck", False)) else 0
        ws.cell(row, 7).value = case_data.get("obs", "")
    row_index[_safe_str(paciente)] = row

    counts = {"MEDIDO": 0, "NO_DISPONIBLE": 0, "PENDIENTE": 0}
    idx_col = 8
    for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS[:26]):
        lm = lm_data[idx] if idx < len(lm_data) else {"status": "OMIT"}
        status = str(lm.get("status", "OMIT")).upper()
        if status == "SET":
            value = "MEDIDO"
            counts["MEDIDO"] += 1
        elif status == "MISSING":
            value = "NO_DISPONIBLE"
            counts["NO_DISPONIBLE"] += 1
        else:
            value = "PENDIENTE"
            counts["PENDIENTE"] += 1
        ws.cell(row, idx_col).value = value
        idx_col += 1

    ws.cell(row, idx_col).value = counts["MEDIDO"]
    ws.cell(row, idx_col + 1).value = counts["NO_DISPONIBLE"]
    ws.cell(row, idx_col + 2).value = counts["PENDIENTE"]
    ws.cell(row, idx_col + 3).value = 1 if counts["PENDIENTE"] == 0 else 0
    ws.cell(row, idx_col + 4).value = round(counts["MEDIDO"] / 26 * 100, 2)
    ws.cell(row, idx_col + 5).value = round((counts["MEDIDO"] + counts["NO_DISPONIBLE"]) / 26 * 100, 2)
    ws.cell(row, idx_col + 6).value = "COMPLETO" if counts["PENDIENTE"] == 0 else "PARCIAL"

    neck_ant_idx = _LM_NAME_TO_IDX.get(_normalize_lm("neck_ant"))
    neck_post_idx = _LM_NAME_TO_IDX.get(_normalize_lm("neck_post"))
    ws.cell(row, idx_col + 7).value = "SET" if neck_ant_idx is not None and neck_ant_idx < len(lm_data) and lm_data[neck_ant_idx].get("status") == "SET" else "OMIT"
    ws.cell(row, idx_col + 8).value = "SET" if neck_post_idx is not None and neck_post_idx < len(lm_data) and lm_data[neck_post_idx].get("status") == "SET" else "OMIT"
    ws.cell(row, idx_col + 9).value = case_data.get("ccu_mm") if case_data.get("has_neck") else None
    ws.cell(row, idx_col + 10).value = timestamp_now()

    new_values = [_safe_str(ws.cell(row, 8 + pos).value) for pos in range(26)]
    _update_resumen_lm_delta(wb, old_values, new_values)


def ensure_resumen_lm_sheet(wb):
    headers = ["Landmark", "N_Medido", "N_NoDisponible", "N_Pendiente", "N_Total", "Porcentaje_Medido"]
    ws = wb[SHEET_RESUMEN_LM] if SHEET_RESUMEN_LM in wb.sheetnames else wb.create_sheet(SHEET_RESUMEN_LM)
    for col, header in enumerate(headers, 1):
        ws.cell(1, col).value = header
    for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS[:26], 2):
        ws.cell(idx, 1).value = lm_id
    _apply_header_style(ws, len(headers))
    return ws


def _write_resumen_lm_sheet(wb):
    source = ensure_disponibilidad_sheet(wb)
    target = ensure_resumen_lm_sheet(wb)
    lm_start = 8
    totals = [{"MEDIDO": 0, "NO_DISPONIBLE": 0, "PENDIENTE": 0} for _ in _OTHER_INDICES]
    for row in range(2, source.max_row + 1):
        for pos in range(26):
            value = _safe_str(source.cell(row, lm_start + pos).value)
            if value in totals[pos]:
                totals[pos][value] += 1
    for pos, counts in enumerate(totals, 2):
        total = sum(counts.values())
        target.cell(pos, 2).value = counts["MEDIDO"]
        target.cell(pos, 3).value = counts["NO_DISPONIBLE"]
        target.cell(pos, 4).value = counts["PENDIENTE"]
        target.cell(pos, 5).value = total
        target.cell(pos, 6).value = round(counts["MEDIDO"] / total * 100, 2) if total else 0.0


def _update_resumen_lm_delta(wb, old_values, new_values):
    """Actualiza Resumen_LM sin recorrer de nuevo todos los pacientes."""
    target = ensure_resumen_lm_sheet(wb)
    for pos in range(26):
        old = old_values[pos] if pos < len(old_values) else ""
        new = new_values[pos] if pos < len(new_values) else ""
        if old == new:
            continue
        row = pos + 2
        counts = {
            "MEDIDO": _safe_int(target.cell(row, 2).value),
            "NO_DISPONIBLE": _safe_int(target.cell(row, 3).value),
            "PENDIENTE": _safe_int(target.cell(row, 4).value),
        }
        if old in counts:
            counts[old] = max(0, counts[old] - 1)
        if new in counts:
            counts[new] += 1
        total = sum(counts.values())
        target.cell(row, 2).value = counts["MEDIDO"]
        target.cell(row, 3).value = counts["NO_DISPONIBLE"]
        target.cell(row, 4).value = counts["PENDIENTE"]
        target.cell(row, 5).value = total
        target.cell(row, 6).value = round(counts["MEDIDO"] / total * 100, 2) if total else 0.0


# =============================================================================
# EXPORTACIÓN A EXCEL Y RESÚMENES DEL CASO
# =============================================================================
# El sistema no solo guarda coordenadas; también registra:
# - datos demográficos del caso;
# - estado de cada landmark;
# - espesor facial por punto;
# - identidad del observador, sesión y réplica;
# - disponibilidad anatómica y proporción de cobertura;
# - medición de cuello (CCU) y estado nutricional.
#
# La serialización del caso a Excel actúa como archivo de trabajo y evidencia del
# estudio, preservando la trazabilidad de cada medición.
# =============================================================================
def _validate_canonical_measurement_context(case_data):
    """Impide que una sesión distinta de NORMAL/S1 sobrescriba hojas canónicas.

    La identidad del registro se define por modo/sesión/observador/réplica.
    ``Timestamp`` es únicamente una bitácora y nunca decide qué medición es principal.
    Una re-revisión de la medición principal debe guardarse otra vez como NORMAL/S1,
    de modo que actualice la misma medición lógica en ``Mediciones_LM`` y los
    resúmenes canónicos.
    """
    mode = _safe_str(case_data.get("measurement_mode", "NORMAL")).upper() or "NORMAL"
    session = _safe_str(case_data.get("session_id", _PRIMARY_MEASUREMENT_SESSION)).upper()
    if mode != "NORMAL":
        raise ValueError("Las hojas canónicas solo pueden escribirse en modo NORMAL.")
    if session != _PRIMARY_MEASUREMENT_SESSION:
        raise ValueError(
            f"La medición principal canónica solo puede guardarse como NORMAL/{_PRIMARY_MEASUREMENT_SESSION}. "
            "Para una segunda medición usa modo ERROR. Si estás re-revisando el dato principal, "
            f"mantén la sesión {_PRIMARY_MEASUREMENT_SESSION}."
        )


def save_landmarks_to_excel(excel_path, case_data, lm_data, unit_scale, mark_processed=False):
    _validate_canonical_measurement_context(case_data)
    global _ACTIVE_SAVE_PERF
    try:
        return _save_landmarks_to_excel_impl(
            excel_path, case_data, lm_data, unit_scale, mark_processed
        )
    finally:
        _ACTIVE_SAVE_PERF = None


def save_error_measurement_to_excel(excel_path, case_data, lm_data):
    """Guarda una sesión de error únicamente en Mediciones_LM.

    No modifica Coordenadas, Resultados, Disponibilidad_LM, Resumen_LM ni el
    estado Procesado del caso principal. Así las remediciones permanecen
    independientes y nunca pisan la observación canónica. La fecha de guardado
    no determina qué sesión es válida.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel no encontrado: {excel_path}")
    with open_workbook(excel_path, save=True) as wb:
        setattr(wb, "_tesis_excel_path", excel_path)
        _append_mediciones_long(wb, case_data, lm_data, excel_path)
    invalidate_excel_cache()


def _save_landmarks_to_excel_impl(excel_path, case_data, lm_data, unit_scale, mark_processed=False):
    """Guarda en una sola apertura todas las hojas relevantes del caso actual."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel no encontrado: {excel_path}")

    global _ACTIVE_SAVE_PERF
    perf = {}
    _ACTIVE_SAVE_PERF = perf
    total_started = time.perf_counter()
    with open_workbook(excel_path, save=True) as wb:
        setattr(wb, "_tesis_excel_path", excel_path)
        ws_cases = wb[SHEET_CASES]
        headers = {ws_cases.cell(1, c).value: c for c in range(1, ws_cases.max_column + 1)}
        paciente = case_data["paciente"]
        ojo = case_data["ojo"]
        case_index = _build_case_row_index(ws_cases, headers)
        row = _find_case_row_prefer_operational(ws_cases, paciente, ojo, headers, case_index)
        if row is None:
            row = ws_cases.max_row + 1
        ws_cases.cell(row, headers.get(COL_PACIENTE, 1)).value = paciente
        ws_cases.cell(row, headers.get(COL_OJO, 2)).value = ojo
        ws_cases.cell(row, headers.get(COL_EDAD, 3)).value = case_data.get("edad", "")
        ws_cases.cell(row, headers.get(COL_SEXO, 4)).value = case_data.get("sexo", "")
        ws_cases.cell(row, headers.get(COL_GRUPO, 5)).value = case_data.get("grupo", "")
        ws_cases.cell(row, headers.get(COL_PESO, 6)).value = case_data.get("peso", "")
        ws_cases.cell(row, headers.get(COL_OBS, 7)).value = case_data.get("obs", "")
        if COL_CUELLO in headers:
            _ws_set_value(ws_cases, row, headers[COL_CUELLO], 1 if bool(case_data.get("has_neck", False)) else 0)
        if COL_FALTAN in headers:
            _ws_set_value(ws_cases, row, headers[COL_FALTAN], 1 if bool(case_data.get("faltan_landmarks", False)) else 0)
        current_processed = False
        if COL_PROCESADO in headers:
            current_processed = bool(_safe_int(ws_cases.cell(row, headers[COL_PROCESADO]).value, 0))
            effective_processed = bool(current_processed or mark_processed)
            _ws_set_value(ws_cases, row, headers[COL_PROCESADO], 1 if effective_processed else 0)
        else:
            effective_processed = bool(mark_processed)

        # Fecha procesado solo cambia al FINALIZAR. Guardar progreso no debe
        # crear ni reemplazar una fecha de finalización.
        if COL_FECHA in headers and bool(mark_processed):
            _ws_set_value(ws_cases, row, headers[COL_FECHA], timestamp_now())

        ws_lm = ensure_landmarks_sheet(wb)
        lm_index = _cached_sheet_index(
            excel_path, ws_lm,
            lambda r: (_safe_str(ws_lm.cell(r, 1).value), _safe_str(ws_lm.cell(r, 2).value)),
        )
        row_lm = _find_patient_row(ws_lm, paciente, ojo, _index=lm_index)
        if row_lm is None:
            row_lm = ws_lm.max_row + 1
        ws_lm.cell(row_lm, 1).value = paciente
        ws_lm.cell(row_lm, 2).value = ojo
        lm_index[_case_pair_key(paciente, ojo)] = row_lm
        for idx, lm in enumerate(lm_data):
            base_col = _LM_DATA_START + idx * _LM_FIELDS_PER_LM
            status = _safe_str(lm.get("status", "OMIT")).upper()
            is_neck = idx in _NECK_INDEX_SET
            ws_lm.cell(row_lm, base_col + 0).value = status
            ws_lm.cell(row_lm, base_col + 1).value = lm.get("hx") if status == "SET" and not is_neck else None
            ws_lm.cell(row_lm, base_col + 2).value = lm.get("hy") if status == "SET" and not is_neck else None
            ws_lm.cell(row_lm, base_col + 3).value = lm.get("hz") if status == "SET" and not is_neck else None
            ws_lm.cell(row_lm, base_col + 4).value = lm.get("x") if status == "SET" else None
            ws_lm.cell(row_lm, base_col + 5).value = lm.get("y") if status == "SET" else None
            ws_lm.cell(row_lm, base_col + 6).value = lm.get("z") if status == "SET" else None
            ws_lm.cell(row_lm, base_col + 7).value = lm.get("thickness") if status == "SET" and not is_neck else None
        end_col = _LM_DATA_START + len(LANDMARKS_ESPERADOS) * _LM_FIELDS_PER_LM
        _ws_set_value(ws_lm, row_lm, end_col + 0, float(unit_scale))
        _ws_set_value(ws_lm, row_lm, end_col + 1, timestamp_now())

        with perf_section("actualizar Resultados", perf):
            _write_resultados_sheet(wb, case_data, lm_data)
        with perf_section("actualizar Mediciones_LM", perf):
            _append_mediciones_long(wb, case_data, lm_data, excel_path)
        case_data["procesado"] = bool(effective_processed)
        with perf_section("actualizar Disponibilidad_LM + Resumen_LM", perf):
            _write_disponibilidad_lm_sheet(wb, case_data, lm_data)

    _ACTIVE_SAVE_PERF = None
    perf["TOTAL"] = time.perf_counter() - total_started
    _print_perf_block("SAVE", perf)
    log.info("Guardado Excel completo para %s / %s", paciente, ojo)


def _write_resultados_sheet(wb, case_data, lm_data):
    """Actualiza la VISTA DERIVADA ``Resultados`` desde la medición NORMAL/S1.

    ``Mediciones_LM`` conserva la identidad explícita de cada medición; ``Resultados``
    es un resumen canónico para consulta y nunca debe usarse para decidir entre
    sesiones, observadores o réplicas. El timestamp es solo informativo.
    """
    ws = ensure_resultados_sheet(wb)

    paciente = case_data["paciente"]
    ojo = case_data["ojo"]

    # Índice O(1) para buscar filas de pacientes
    excel_path = getattr(wb, "_tesis_excel_path", None)
    res_index = (_cached_sheet_index(
        excel_path, ws,
        lambda r: (_safe_str(ws.cell(r, 1).value), _safe_str(ws.cell(r, 2).value)),
    ) if excel_path else _build_patient_row_index(ws))

    # Buscar fila existente o crear nueva
    row = _find_patient_row(ws, paciente, ojo, _index=res_index)
    if row is None:
        row = ws.max_row + 1

    ws.cell(row, 1).value = paciente
    ws.cell(row, 2).value = ojo
    if isinstance(res_index, dict):
        res_index[_case_pair_key(paciente, ojo)] = row
    ws.cell(row, 3).value = case_data["edad"]
    ws.cell(row, 4).value = case_data["sexo"]

    # Grosor por landmark (1 columna por landmark)
    for facial_pos, idx in enumerate(_OTHER_INDICES):
        col = _RES_DATA_START + facial_pos
        lm = lm_data[idx]
        if lm["status"] == "SET" and lm["thickness"] is not None and lm["thickness"] > 0:
            ws.cell(row, col).value = lm["thickness"]
        else:
            ws.cell(row, col).value = None

    # CCU y estado nutricional al final
    end_col = _RES_DATA_START + len(_OTHER_INDICES)
    if case_data.get("has_neck", False):
        ccu_val = round(float(case_data.get("ccu_mm", 0.0)), 2)
        estado_val = estado_nutricional(case_data.get("sexo", ""), case_data.get("edad", ""), ccu_val)
    else:
        ccu_val = None
        estado_val = "No aplicable"

    ws.cell(row, end_col + 0).value = ccu_val
    ws.cell(row, end_col + 1).value = estado_val
    ws.cell(row, end_col + 2).value = timestamp_now()


def load_landmarks_from_excel(case):
    """Lee landmarks en formato horizontal (1 fila por paciente)."""
    excel_path = get_excel_path()
    if not os.path.exists(excel_path):
        return {}

    t0 = time.perf_counter()
    with open_workbook(excel_path, read_only=True, data_only=True) as wb:
        if SHEET_LANDMARKS not in wb.sheetnames:
            return {}

        ws = wb[SHEET_LANDMARKS]

        # Detectar formato leyendo primera fila con iter_rows
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row and len(header_row) > 2 and _safe_str(header_row[2]) == "Landmark":
            return {}  # formato viejo vertical, no soportado

        # Buscar fila del paciente usando streaming (iter_rows)
        target_pac = str(case.paciente)
        target_ojo = str(case.ojo)
        target_row = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            if str(row[0]) == target_pac and str(row[1]) == target_ojo:
                target_row = row
                break

        if target_row is None:
            return {}

        row_len = len(target_row)
        data = {}
        for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS):
            base = _LM_DATA_START + idx * _LM_FIELDS_PER_LM - 1  # 0-based
            if base >= row_len:
                continue
            status = target_row[base]
            if status is None:
                continue
            data[lm_id] = {
                "status": str(status),
                "hx": _safe_float(target_row[base + 1], 0.0) if base + 1 < row_len else 0.0,
                "hy": _safe_float(target_row[base + 2], 0.0) if base + 2 < row_len else 0.0,
                "hz": _safe_float(target_row[base + 3], 0.0) if base + 3 < row_len else 0.0,
                "sx": _safe_float(target_row[base + 4], 0.0) if base + 4 < row_len else 0.0,
                "sy": _safe_float(target_row[base + 5], 0.0) if base + 5 < row_len else 0.0,
                "sz": _safe_float(target_row[base + 6], 0.0) if base + 6 < row_len else 0.0,
                "thickness": _safe_float(target_row[base + 7], 0.0) if base + 7 < row_len else 0.0,
            }

    log.debug("load_landmarks_from_excel para %s/%s en %.2fs",
              case.paciente, case.ojo, time.perf_counter() - t0)
    return data


def _error_progress_key_values(case_id, observer, session, replica, cohort):
    return (
        _safe_str(case_id),
        _safe_str(observer).upper(),
        _safe_str(session).upper(),
        int(replica or 1),
        _safe_str(cohort),
    )


def _error_progress_key(case, state):
    return _error_progress_key_values(
        getattr(case, "paciente", ""),
        getattr(state, "observer_id", "OBS1"),
        getattr(state, "session_id", "S1"),
        getattr(state, "replicate_id", 1),
        getattr(state, "error_cohort", _DEFAULT_ERROR_COHORT),
    )


def _rebuild_error_progress_cache_from_ws(ws):
    """Reconstruye el progreso de sesiones ERROR desde Mediciones_LM."""
    global _ERROR_PROGRESS_CACHE
    cache = {}
    if ws is None or ws.max_row < 2:
        _ERROR_PROGRESS_CACHE = cache
        return cache

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_row:
        _ERROR_PROGRESS_CACHE = cache
        return cache
    headers = {_safe_str(value): index for index, value in enumerate(header_row)}
    required = {"CaseID", "Observador", "Sesion", "Replica", "Landmark", "Status"}
    if not required.issubset(headers):
        _ERROR_PROGRESS_CACHE = cache
        return cache

    mode_col = headers.get("ModoMedicion")
    cohort_col = headers.get("CohorteError")
    ccu_col = headers.get("CCU_mm")
    for row in ws.iter_rows(min_row=2, values_only=True):
        mode = _safe_str(row[mode_col]).upper() if mode_col is not None and mode_col < len(row) else "NORMAL"
        if mode != "ERROR":
            continue
        cohort = _safe_str(row[cohort_col]) if cohort_col is not None and cohort_col < len(row) else ""
        key = _error_progress_key_values(
            row[headers["CaseID"]],
            row[headers["Observador"]],
            row[headers["Sesion"]],
            _safe_int(row[headers["Replica"]], 1),
            cohort,
        )
        entry = cache.setdefault(key, {"statuses": {}, "ccu_mm": 0.0})
        lm_id = _safe_str(row[headers["Landmark"]])
        if lm_id:
            entry["statuses"][lm_id] = _safe_str(row[headers["Status"]]).upper() or "OMIT"
        if ccu_col is not None and ccu_col < len(row):
            ccu = _safe_float(row[ccu_col], 0.0)
            if ccu > 0:
                entry["ccu_mm"] = ccu
    _ERROR_PROGRESS_CACHE = cache
    return cache


def _error_session_counts(state):
    cohort = _safe_str(getattr(state, "error_cohort", ""))
    cases = [case for case in getattr(state, "cases", [])
             if bool(getattr(case, "error_study", False))
             and (not cohort or _safe_str(getattr(case, "error_cohort", "")) == cohort)]
    complete = sum(1 for case in cases if _error_session_complete(case, state))
    return len(cases), complete, max(0, len(cases) - complete)


def _error_session_complete(case, state):
    entry = _ERROR_PROGRESS_CACHE.get(_error_progress_key(case, state), {})
    statuses = entry.get("statuses", {}) if isinstance(entry, dict) else {}
    for idx in _OTHER_INDICES:
        lm_id = LANDMARKS_ESPERADOS[idx][0]
        if _safe_str(statuses.get(lm_id, "OMIT")).upper() not in {"SET", "MISSING"}:
            return False
    if bool(getattr(state, "error_include_neck", True)):
        if not bool(getattr(case, "has_neck", False)):
            return False
        for idx in _NECK_INDICES:
            lm_id = LANDMARKS_ESPERADOS[idx][0]
            if _safe_str(statuses.get(lm_id, "OMIT")).upper() != "SET":
                return False
        if _safe_float(entry.get("ccu_mm", 0.0), 0.0) <= 0:
            return False
    return True


def _load_error_measurement_session_from_excel(case, state):
    """Lee solo la sesión ERROR activa; nunca cae hacia Coordenadas canónicas."""
    excel_path = get_excel_path()
    if not os.path.exists(excel_path):
        return {}, 0.0
    target_case = _safe_str(case.paciente)
    target_observer = _safe_str(state.observer_id).upper()
    target_session = _safe_str(state.session_id).upper()
    target_replica = int(state.replicate_id or 1)
    target_cohort = _safe_str(state.error_cohort)
    data = {}
    ccu_mm = 0.0

    with open_workbook(excel_path, read_only=True, data_only=True) as wb:
        if SHEET_MEDICIONES not in wb.sheetnames:
            return {}, 0.0
        ws = wb[SHEET_MEDICIONES]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            return {}, 0.0
        headers = {_safe_str(v): i for i, v in enumerate(header_row) if v is not None}
        required = ["CaseID", "Observador", "Sesion", "Replica", "Landmark", "Status"]
        if any(name not in headers for name in required):
            return {}, 0.0

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            if _safe_str(row[headers["CaseID"]]) != target_case:
                continue
            if _safe_str(row[headers["Observador"]]).upper() != target_observer:
                continue
            if _safe_str(row[headers["Sesion"]]).upper() != target_session:
                continue
            if _safe_int(row[headers["Replica"]], 1) != target_replica:
                continue
            mode = _safe_str(row[headers["ModoMedicion"]]).upper() if "ModoMedicion" in headers else "NORMAL"
            cohort = _safe_str(row[headers["CohorteError"]]) if "CohorteError" in headers else ""
            if mode != "ERROR" or cohort != target_cohort:
                continue

            lm_id = _safe_str(row[headers["Landmark"]])
            if not lm_id:
                continue
            status = _safe_str(row[headers["Status"]]).upper() or "OMIT"
            def val(name):
                return _safe_float(row[headers[name]], 0.0) if name in headers and headers[name] < len(row) else 0.0
            data[lm_id] = {
                "status": status,
                "hx": val("HX"), "hy": val("HY"), "hz": val("HZ"),
                "sx": val("SX"), "sy": val("SY"), "sz": val("SZ"),
                "thickness": val("Thickness_mm"),
            }
            if "CCU_mm" in headers and headers["CCU_mm"] < len(row):
                ccu = _safe_float(row[headers["CCU_mm"]], 0.0)
                if ccu > 0:
                    ccu_mm = ccu
    return data, ccu_mm


def _load_active_measurement_data(case, state):
    if _safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() == "ERROR":
        data, ccu = _load_error_measurement_session_from_excel(case, state)
        case.ccu_mm = float(ccu or 0.0)
        case.estado_nutricional = (
            estado_nutricional(case.sexo, case.edad, case.ccu_mm)
            if case.ccu_mm > 0 else ""
        )
        return data
    case.ccu_mm = float(getattr(case, "baseline_ccu_mm", 0.0) or 0.0)
    case.estado_nutricional = _safe_str(getattr(case, "baseline_estado_nutricional", ""))
    return load_landmarks_from_excel(case)


def sync_landmarks_state_from_excel(case, state):
    """Sincroniza el caso con la fuente de medición activa.

    NORMAL -> hoja Coordenadas canónica.
    ERROR  -> Mediciones_LM del observador/sesión/réplica/cohorte actual.

    En modo ERROR nunca se reutilizan las coordenadas de la medición principal
    si la sesión activa todavía está vacía.
    """
    _init_landmarks_for_case(state)
    data = _load_active_measurement_data(case, state)
    if not data:
        _update_lm_counts(state)
        return 0

    loaded_set = 0
    for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS):
        d = data.get(lm_id)
        if not d:
            continue

        lm = state.landmarks[idx]
        lm.status = d.get("status", "OMIT")
        lm.hx = d.get("hx", 0.0)
        lm.hy = d.get("hy", 0.0)
        lm.hz = d.get("hz", 0.0)
        lm.x = d.get("sx", 0.0)
        lm.y = d.get("sy", 0.0)
        lm.z = d.get("sz", 0.0)
        lm.thickness = d.get("thickness", 0.0)

        if lm.status == "SET":
            loaded_set += 1

    _update_lm_counts(state)
    return loaded_set


_COORDINATE_FRAME = "CASE_FRAME_V1"
_COORDINATE_COMPATIBLE_MM = 0.5
_COORDINATE_WARNING_MM = 2.0
_FRAME_VALIDATION_CACHE = {}


def _frame_validation_key(case, skull, soft):
    paths = (getattr(case, "ruta_hueso", ""), getattr(case, "ruta_blando", ""))
    file_info = []
    for path in paths:
        try:
            stat = os.stat(path)
            file_info.append((path, stat.st_mtime_ns, stat.st_size))
        except OSError:
            file_info.append((path, None, None))
    return (_current_case_key(case), tuple(file_info), skull.name, soft.name)


def _legacy_orientation_matrix():
    return Matrix.Rotation(math.radians(-90.0), 4, 'Y') @ Matrix.Rotation(math.radians(90.0), 4, 'X')


def _build_mesh_bvh_for_orientation(obj, orientation="KEEP"):
    depsgraph = bpy.context.evaluated_depsgraph_get()
    obj_eval = obj.evaluated_get(depsgraph)
    mesh_eval = obj_eval.to_mesh()
    try:
        transform = _legacy_orientation_matrix() if orientation == "LEGACY" else Matrix.Identity(4)
        vertices = [obj_eval.matrix_world @ transform @ vertex.co for vertex in mesh_eval.vertices]
        polygons = [polygon.vertices[:] for polygon in mesh_eval.polygons]
        return BVHTree.FromPolygons(vertices, polygons)
    finally:
        obj_eval.to_mesh_clear()


def _saved_landmark_points(data):
    points = []
    for idx, (lm_id, _soft_id, _label) in enumerate(LANDMARKS_ESPERADOS):
        if idx in _NECK_INDEX_SET:
            continue
        saved = data.get(lm_id)
        if not saved or str(saved.get("status", "")).upper() != "SET":
            continue
        points.append((lm_id, Vector((saved.get("hx", 0.0), saved.get("hy", 0.0), saved.get("hz", 0.0))),
                      Vector((saved.get("sx", 0.0), saved.get("sy", 0.0), saved.get("sz", 0.0)))))
    return points


def _median_saved_surface_errors(data, skull, soft, orientation="KEEP"):
    skull_bvh = _BVH_CACHE.get("skull") if orientation == "KEEP" else None
    soft_bvh = _BVH_CACHE.get("soft") if orientation == "KEEP" else None
    if skull_bvh is None or soft_bvh is None:
        skull_bvh = _build_mesh_bvh_for_orientation(skull, orientation)
        soft_bvh = _build_mesh_bvh_for_orientation(soft, orientation)
    hard_errors = []
    soft_errors = []
    for _lm_id, hard_point, soft_point in _saved_landmark_points(data):
        hard_hit = skull_bvh.find_nearest(hard_point)
        soft_hit = soft_bvh.find_nearest(soft_point)
        if hard_hit[0] is not None:
            hard_errors.append(float(hard_hit[3]))
        if soft_hit[0] is not None:
            soft_errors.append(float(soft_hit[3]))
    return (
        statistics.median(hard_errors) if hard_errors else float("inf"),
        statistics.median(soft_errors) if soft_errors else float("inf"),
    )


def _median_saved_surface_errors_from_keep(data, skull, soft, skull_bvh, soft_bvh, orientation):
    """Evalúa KEEP o LEGACY usando el BVH KEEP y una transformación rígida."""
    if orientation == "KEEP":
        return _median_saved_surface_errors(data, skull, soft, "KEEP")
    skull_transform = skull.matrix_world @ _legacy_orientation_matrix() @ skull.matrix_world.inverted_safe()
    soft_transform = soft.matrix_world @ _legacy_orientation_matrix() @ soft.matrix_world.inverted_safe()
    skull_inverse = skull_transform.inverted_safe()
    soft_inverse = soft_transform.inverted_safe()
    hard_errors = []
    soft_errors = []
    for _lm_id, hard_point, soft_point in _saved_landmark_points(data):
        hard_hit = skull_bvh.find_nearest(skull_inverse @ hard_point)
        soft_hit = soft_bvh.find_nearest(soft_inverse @ soft_point)
        if hard_hit[0] is not None:
            hard_errors.append(float(hard_hit[3]))
        if soft_hit[0] is not None:
            soft_errors.append(float(soft_hit[3]))
    return (
        statistics.median(hard_errors) if hard_errors else float("inf"),
        statistics.median(soft_errors) if soft_errors else float("inf"),
    )


def _format_frame_diagnostic(patient, results, best, valid):
    patient = _safe_str(patient)
    lines = [
        "Coordenadas históricas no compatibles." if not valid else "Coordenadas históricas compatibles.",
        f"Paciente: {patient}",
        "",
        "KEEP:",
        f"HARD mediana = {results['KEEP']['hard_median']:.2f} mm",
        f"SOFT mediana = {results['KEEP']['soft_median']:.2f} mm",
        "",
        "LEGACY:",
        f"HARD mediana = {results['LEGACY']['hard_median']:.2f} mm",
        f"SOFT mediana = {results['LEGACY']['soft_median']:.2f} mm",
        "",
        f"Mejor candidato: {best}",
        f"Error máximo mediano: {max(results[best].values()):.2f} mm",
        "No se modificó ninguna coordenada ni medición.",
    ]
    return "\n".join(lines)


# =============================================================================
# VALIDACIÓN DE COORDENADAS HISTÓRICAS
# =============================================================================
# Antes de aceptar un caso como compatible, el programa compara las coordenadas
# guardadas con la geometría actual de las mallas. Esto evita trabajar con
# archivos provenientes de un sistema de referencia distinto o con un cambio de
# orientación no controlado.
#
# La validación no reemplaza la medición ni manipula la malla; simplemente
# decide si el caso puede continuar sin riesgo de desalineación.
# =============================================================================
def validate_saved_landmarks_against_meshes(data, skull, soft, cache_key=None, patient="Caso actual"):
    """Selecciona el marco compatible sin modificar mallas ni coordenadas."""
    if cache_key is not None and cache_key in _FRAME_VALIDATION_CACHE:
        return _FRAME_VALIDATION_CACHE[cache_key]
    if not _saved_landmark_points(data):
        empty = {"KEEP": {"hard_median": 0.0, "soft_median": 0.0},
             "LEGACY": {"hard_median": 0.0, "soft_median": 0.0}}
        result = ("KEEP", empty, True, "No hay landmarks SET guardados para validar.")
        if cache_key is not None:
            _FRAME_VALIDATION_CACHE[cache_key] = result
        return result

    keep_bvh = _BVH_CACHE.get("skull")
    soft_bvh = _BVH_CACHE.get("soft")
    if keep_bvh is None or soft_bvh is None:
        keep_bvh = build_bvh_world(skull, bpy.context.evaluated_depsgraph_get())
        soft_bvh = build_bvh_world(soft, bpy.context.evaluated_depsgraph_get())
        _BVH_CACHE["skull"] = keep_bvh
        _BVH_CACHE["soft"] = soft_bvh
        _BVH_CACHE["skull_name"] = skull.name
        _BVH_CACHE["soft_name"] = soft.name
        _BVH_CACHE["skull_sig"] = _bvh_object_signature(skull)
        _BVH_CACHE["soft_sig"] = _bvh_object_signature(soft)
    keep_errors = _median_saved_surface_errors_from_keep(
        data, skull, soft, keep_bvh, soft_bvh, "KEEP"
    )
    raw_results = {"KEEP": keep_errors}
    if max(keep_errors) > _COORDINATE_COMPATIBLE_MM:
        raw_results["LEGACY"] = _median_saved_surface_errors_from_keep(
            data, skull, soft, keep_bvh, soft_bvh, "LEGACY"
        )
    else:
        raw_results["LEGACY"] = (float("inf"), float("inf"))
    results = {
        name: {"hard_median": values[0], "soft_median": values[1]}
        for name, values in raw_results.items()
    }
    orientation = min(results, key=lambda key: max(results[key].values()))
    hard_error = results[orientation]["hard_median"]
    soft_error = results[orientation]["soft_median"]
    worst_error = max(hard_error, soft_error)
    if worst_error > _COORDINATE_WARNING_MM:
        message = _format_frame_diagnostic(patient, results, orientation, False)
        print("[TESIS FRAME DIAGNOSTIC]")
        print(f"Paciente: {message.splitlines()[1].replace('Paciente: ', '')}")
        print(f"KEEP HARD: {results['KEEP']['hard_median']:.2f} mm")
        print(f"KEEP SOFT: {results['KEEP']['soft_median']:.2f} mm")
        print(f"LEGACY HARD: {results['LEGACY']['hard_median']:.2f} mm")
        print(f"LEGACY SOFT: {results['LEGACY']['soft_median']:.2f} mm")
        print(f"BEST: {orientation}")
        print("VALID: False")
        result = (None, results, False, message)
        if cache_key is not None:
            _FRAME_VALIDATION_CACHE[cache_key] = result
        return result
    if worst_error > _COORDINATE_COMPATIBLE_MM:
        result = (orientation, results, True, (
            f"Marco {orientation}: error mediano HARD {hard_error:.2f} mm / SOFT {soft_error:.2f} mm; revisar."
        ))
    else:
        result = (orientation, results, True, f"Marco {orientation} compatible.")
    if cache_key is not None:
        _FRAME_VALIDATION_CACHE[cache_key] = result
    return result


def diagnose_coordinate_frame(context):
    """Imprime un diagnóstico KEEP/LEGACY sin modificar estado, mallas ni Excel."""
    state = context.scene.tesis_state
    if not state.cases or state.active_case >= len(state.cases):
        raise ValueError("No hay caso activo.")
    case = state.cases[state.active_case]
    skull, soft = obtener_mallas_paciente(quiet=True)
    if skull is None or soft is None:
        raise ValueError("Importa primero SKULL y SOFT para diagnosticar.")

    data = load_landmarks_from_excel(case)
    bvh_keep = (_build_mesh_bvh_for_orientation(skull, "KEEP"),
                _build_mesh_bvh_for_orientation(soft, "KEEP"))
    bvh_legacy = (_build_mesh_bvh_for_orientation(skull, "LEGACY"),
                  _build_mesh_bvh_for_orientation(soft, "LEGACY"))
    rows = []
    for lm_id, hard_point, soft_point in _saved_landmark_points(data):
        values = []
        for skull_bvh, soft_bvh in (bvh_keep, bvh_legacy):
            hard_hit = skull_bvh.find_nearest(hard_point)
            soft_hit = soft_bvh.find_nearest(soft_point)
            values.extend([
                float(hard_hit[3]) if hard_hit[0] is not None else float("inf"),
                float(soft_hit[3]) if soft_hit[0] is not None else float("inf"),
            ])
        rows.append((max(values), lm_id, values))
    rows.sort(key=lambda row: row[0], reverse=True)

    def file_info(path):
        try:
            stat = os.stat(path)
            return (os.path.basename(path), stat.st_size, stat.st_mtime)
        except OSError:
            return (os.path.basename(path), None, None)

    print("[TESIS FRAME DIAGNOSTIC]")
    print(f"Paciente: {case.paciente}")
    print(f"Ruta SKULL: {case.ruta_hueso} | archivo/tamaño/mtime: {file_info(case.ruta_hueso)}")
    print(f"Ruta SOFT: {case.ruta_blando} | archivo/tamaño/mtime: {file_info(case.ruta_blando)}")
    print(f"SKULL vertices/polígonos: {len(skull.data.vertices)}/{len(skull.data.polygons)}")
    print(f"SOFT vertices/polígonos: {len(soft.data.vertices)}/{len(soft.data.polygons)}")
    print("LM       KEEP_H  KEEP_S  LEG_H   LEG_S")
    for _worst, lm_id, values in rows:
        print(f"{lm_id:<8} {values[0]:7.2f} {values[1]:7.2f} {values[2]:7.2f} {values[3]:7.2f}")

    keep = _median_saved_surface_errors(data, skull, soft, "KEEP")
    legacy = _median_saved_surface_errors(data, skull, soft, "LEGACY")
    best = "KEEP" if max(keep) <= max(legacy) else "LEGACY"
    messages = []
    if max(keep) > _COORDINATE_WARNING_MM and max(legacy) > _COORDINATE_WARNING_MM:
        messages.append("Posible transformación global/origen diferente.")
    if max(legacy) + _COORDINATE_COMPATIBLE_MM < max(keep):
        messages.append("Posible orientación histórica LEGACY.")
    if keep[0] <= _COORDINATE_COMPATIBLE_MM < keep[1]:
        messages.append("SKULL compatible pero SOFT no compatible: revisar ruta/malla de tejido blando.")
    if keep[1] <= _COORDINATE_COMPATIBLE_MM < keep[0]:
        messages.append("SOFT compatible pero SKULL no compatible: revisar ruta/malla ósea.")
    for message in messages:
        print(message)
    return rows


class TESIS_OT_DiagnoseCoordinateFrame(Operator):
    bl_idname = "tesis.diagnose_coordinate_frame"
    bl_label = "Diagnosticar coordenadas"
    bl_description = "Analiza KEEP y LEGACY sin modificar mediciones"

    def execute(self, context):
        try:
            diagnose_coordinate_frame(context)
        except Exception as exc:
            self.report({'ERROR'}, str(exc))
            return {'CANCELLED'}
        self.report({'INFO'}, "Diagnóstico de coordenadas impreso en consola.")
        return {'FINISHED'}


def rebuild_landmarks_from_data(context, case, state, data=None):
    """Reconstruye objetos 3D de landmarks desde datos guardados en Excel.
    
    Returns:
        int: Número de landmarks reconstruidos (status=SET)
    """
    _OBJ_NAME_TO_LM.clear()
    clear_objects_prefix([f"{i:02d}_" for i in range(len(LANDMARKS_ESPERADOS))])

    if data is None:
        data = _load_active_measurement_data(case, state)

    if not data:
        return 0

    cols = ensure_tesis_collections()
    reconstructed_count = 0
    inconsistent_count = 0
    case.reconstruction_warning = ""

    for idx, (lm_id, _soft_id, lm_label) in enumerate(LANDMARKS_ESPERADOS):

        if lm_id not in data:
            continue

        d = data[lm_id]

        lm = state.landmarks[idx]
        lm.status = d["status"]
        lm.hx, lm.hy, lm.hz = d["hx"], d["hy"], d["hz"]
        lm.x,  lm.y,  lm.z  = d["sx"], d["sy"], d["sz"]
        lm.thickness = d["thickness"]

        if lm.status != "SET":
            continue

        p_hard = Vector((lm.hx, lm.hy, lm.hz))
        p_soft = Vector((lm.x, lm.y, lm.z))

        lm_normalized = normalize_landmark_name(lm_id)
        is_neck = lm_normalized in _NECK_IDS

        # Cuello no posee HARD ni espesor facial. No comparar SOFT contra el origen.
        if not is_neck:
            geometry_thickness = (p_soft - p_hard).length
            if abs(geometry_thickness - float(lm.thickness)) > 0.05:
                inconsistent_count += 1
                if not case.reconstruction_warning:
                    case.reconstruction_warning = (
                        f"El landmark {lm_id} reconstruido no coincide con su espesor guardado: "
                        f"Excel {float(lm.thickness):.2f} / geometría {geometry_thickness:.2f}. "
                        "No se modificó el Excel."
                    )
                log.warning(
                    "%s reconstruido no coincide con espesor guardado: Excel %.4f / geometría %.4f",
                    lm_id, float(lm.thickness), geometry_thickness,
                )

        nice_name = f"{idx:02d}_{lm_id}"

        if is_neck:
            # CUELLO: solo esfera verde en tejido blando
            soft_obj = create_landmark_sphere(
                name=f"{nice_name}_SOFT",
                location=p_soft,
                radius_mm=1.5,
                color_rgba=(0, 1, 0.3, 1),
                collection_name=cols["soft"].name
            )
            soft_obj.hide_select = False
            soft_obj.lock_location = (False, False, False)
            lm.soft_obj_name = soft_obj.name
            lm.hard_obj_name = ""
            lm.line_obj_name = ""
            _register_lm_objs(idx, soft_obj.name)
        else:
            # NORMAL: hard (rojo) + soft (azul) + línea
            hard_obj = create_landmark_sphere(
                name=f"{nice_name}_HARD",
                location=p_hard,
                radius_mm=1.0,
                color_rgba=(1, 0, 0, 1),
                collection_name=cols["hard"].name
            )
            hard_obj.hide_select = False
            hard_obj.lock_location = (False, False, False)
            lm.hard_obj_name = hard_obj.name

            soft_obj = create_landmark_sphere(
                name=f"{nice_name}_SOFT",
                location=p_soft,
                radius_mm=1.0,
                color_rgba=(0, 0.6, 1, 1),
                collection_name=cols["soft"].name
            )
            soft_obj.hide_select = False
            soft_obj.lock_location = (False, False, False)
            lm.soft_obj_name = soft_obj.name

            line_obj = create_line_object(
                name=f"{nice_name}_LINE",
                p0=p_hard,
                p1=p_soft,
                hard_obj=hard_obj,
                soft_obj=soft_obj,
            )
            lm.line_obj_name = line_obj.name
            line_obj["tesis_lm_index"] = int(idx)
            line_obj["tesis_lm_role"] = "LINE"
            _register_lm_objs(idx, soft_obj.name, hard_obj.name)
        
        reconstructed_count += 1
    
    _update_lm_counts(state)
    if inconsistent_count:
        log.warning("%d landmarks reconstruidos con diferencia de espesor superior a 0.05 mm.", inconsistent_count)
    return reconstructed_count

def get_or_create_collection(name: str, parent: bpy.types.Collection | None = None) -> bpy.types.Collection:
    col = bpy.data.collections.get(name)
    if col is None:
        col = bpy.data.collections.new(name)
    if parent is None:
        parent = bpy.context.scene.collection
    assert parent is not None
    # link seguro al parent (comparar objeto, no nombre)
    if col not in parent.children.values():
        parent.children.link(col)
    return col

def ensure_tesis_collections():
    
    root = get_or_create_collection("TESIS_ROOT", bpy.context.scene.collection)
    meshes = get_or_create_collection("TESIS_MESHES", root)
    ccu = get_or_create_collection("TESIS_CCU", root)

    lm_root = get_or_create_collection("TESIS_LANDMARKS", root)
    lm_hard = get_or_create_collection("LM_HARD", lm_root)
    lm_soft = get_or_create_collection("LM_SOFT", lm_root)
    lm_lines = get_or_create_collection("LM_LINES", lm_root)

    return {
        "root": root,
        "meshes": meshes,
        "ccu": ccu,
        "lm_root": lm_root,
        "hard": lm_hard,
        "soft": lm_soft,
        "lines": lm_lines,
    }

def move_object_to_collection(obj: bpy.types.Object, target: bpy.types.Collection):
    """Saca el objeto de TODAS sus colecciones actuales y lo deja solo en target."""
    for col in list(obj.users_collection):
        try:
            col.objects.unlink(obj)
        except Exception:
            pass
    if obj not in target.objects.values():
        target.objects.link(obj)


def _current_case_key(case) -> str:
    return f"{_safe_str(getattr(case, 'paciente', ''))}__{_safe_str(getattr(case, 'ojo', ''))}"


def _pick_imported_mesh(candidates, filepath_hint=""):
    """Elige el objeto malla más probable tras una importación PLY."""
    meshes = [obj for obj in candidates if getattr(obj, "type", None) == 'MESH']
    if not meshes:
        return None

    hint = _safe_str(filepath_hint)
    if hint:
        base = os.path.splitext(os.path.basename(hint))[0].lower()
        exact = [obj for obj in meshes if obj.name.lower() == base]
        if exact:
            return exact[0]
        partial = [obj for obj in meshes if base in obj.name.lower()]
        if partial:
            meshes = partial

    return max(meshes, key=lambda o: len(getattr(getattr(o, "data", None), "vertices", [])))


def _import_ply_and_get_object(context, filepath):
    before = {obj.as_pointer() for obj in bpy.data.objects}
    previous_active = context.view_layer.objects.active
    previous_active_pointer = previous_active.as_pointer() if previous_active is not None else None
    bpy.ops.wm.ply_import(filepath=filepath)
    new_objects = [obj for obj in bpy.data.objects if obj.as_pointer() not in before]
    active = context.view_layer.objects.active
    if (active is not None and active.type == 'MESH'
            and active.as_pointer() not in before
            and active.as_pointer() != previous_active_pointer):
        return active
    return _pick_imported_mesh(new_objects, filepath)


def _tag_case_mesh(obj, role: str, case):
    if obj is None:
        return
    obj["tesis_role"] = role.upper()
    obj["tesis_case_key"] = _current_case_key(case)
    obj.data["tesis_owned"] = True
    obj.data["tesis_case_key"] = _current_case_key(case)
    obj.data["tesis_role"] = role.upper()
    obj.lock_location = (True, True, True)
    obj.lock_rotation = (True, True, True)
    obj.lock_scale = (True, True, True)
    if role.upper() in {"SKULL", "SOFT"}:
        obj.hide_select = True


def _ensure_object_visible(obj):
    if obj is None:
        return
    try:
        obj.hide_set(False)
    except Exception:
        pass
    try:
        obj.hide_viewport = False
        obj.hide_render = False
    except Exception:
        pass


def _clear_case_visual_objects():
    """Elimina objetos visuales del caso activo sin tocar metadatos del caso ni Excel."""
    invalidate_mallas_cache()
    prefixes = ["SKULL_", "SOFT_", "CCU_"] + [f"{i:02d}_" for i in range(len(LANDMARKS_ESPERADOS))]
    state = getattr(getattr(bpy.context, "scene", None), "tesis_state", None)
    if not state or not getattr(state, "cases", None):
        clear_objects_prefix(prefixes)
        return
    try:
        case = state.cases[state.active_case]
        case_key = _current_case_key(case)
    except Exception:
        return

    collections = ("TESIS_MESHES", "LM_HARD", "LM_SOFT", "LM_LINES", "TESIS_CCU")
    candidates = []
    for collection_name in collections:
        collection = bpy.data.collections.get(collection_name)
        if collection is None:
            continue
        candidates.extend(collection.objects)
    targets = [
        obj for obj in set(candidates)
        if any(obj.name.startswith(prefix) for prefix in prefixes)
        or _safe_str(obj.get("tesis_case_key")) == case_key
    ]
    if not targets:
        targets = [
            obj for obj in bpy.data.objects
            if any(obj.name.startswith(prefix) for prefix in prefixes)
            or _safe_str(obj.get("tesis_case_key")) == case_key
        ]
    for obj in targets:
        _safe_remove_object(obj)
    _cleanup_owned_orphans()
    _memory_diagnostic(case_key)


def delete_landmark_complete(state, idx: int) -> bool:
    """Elimina el landmark completo y deja su estado lógico en OMIT."""
    if state is None or not state.landmarks:
        return False
    if idx < 0 or idx >= len(state.landmarks):
        return False

    lm = state.landmarks[idx]
    for attr in ("hard_obj_name", "soft_obj_name", "line_obj_name"):
        name = _safe_str(getattr(lm, attr, ""))
        if name:
            obj = bpy.data.objects.get(name)
            if obj is not None:
                _safe_remove_object(obj)
    clear_landmark_objects(idx)
    _unregister_lm_objs(idx)
    _OBJ_NAME_TO_LM.pop(f"{idx}:SOFT", None)
    _OBJ_NAME_TO_LM.pop(f"{idx}:HARD", None)
    lm.status = "OMIT"
    lm.hx = lm.hy = lm.hz = 0.0
    lm.x = lm.y = lm.z = 0.0
    lm.thickness = 0.0
    lm.hard_obj_name = ""
    lm.soft_obj_name = ""
    lm.line_obj_name = ""
    state.dirty = True
    _update_lm_counts(state)
    return True

# =====================================================
# UI LIST

# =====================================================



def _draw_utilidades_minimal(layout):
    box = layout.box()
    box.label(text="UTILIDADES")

    r1 = box.row(align=True)
    r1.operator("tesis.refresh_bvh", text="Actualizar BVH", icon='FILE_REFRESH')
    r1.operator("tesis.load_saved_landmarks", text="Reconstruir LM", icon='RECOVER_LAST')
    r1.operator("tesis.import_meshes", text="Importar mallas", icon='IMPORT')


def _ui_lm_icon(lm_item):
    if lm_item.status == "SET":
        return 'CHECKMARK'
    if lm_item.status == "MISSING":
        return 'CANCEL'
    return 'CHECKBOX_DEHLT'


def _ui_landmark_label(idx):
    labels = (
        "1.g", "2.n", "3.mn", "4.rhi", "5.sn", "6.mp",
        "7.ls", "8.li", "9.mls", "10.pg", "11.gn", "12.m",
        "13.msoI", "14.msoD", "15.mioI", "16.mioD", "17.acpI", "18.acpD",
        "19.goI", "20.goD", "21.zyI", "22.zyD", "23.mrI", "24.mrD",
        "25.mmbI", "26.mmbD",
    )
    return labels[idx] if 0 <= idx < len(labels) else ""


def _ui_landmark_value(lm):
    if lm.status == "SET" and lm.thickness > 0:
        return f"{lm.thickness:.1f}"
    return "—"


def _ui_lm_mm_text(lm_item):
    return _ui_landmark_value(lm_item)


def _draw_landmark_entry(layout, state, idx, meshes_ok=True):
    """Dibuja una entrada compacta: [número.nombre][valor].

    La etiqueta completa funciona como botón para seleccionar/marcar el landmark.
    El valor ocupa una fracción pequeña y estable del ancho de la celda.
    """
    if idx < 0 or idx >= len(state.landmarks):
        return

    lm = state.landmarks[idx]
    row = layout.row(align=True)
    row.scale_y = 0.82
    if _landmark_ui_warn(state, idx):
        row.alert = True

    split = row.split(factor=0.72, align=True)

    mark = split.row(align=True)
    mark.enabled = meshes_ok
    op = mark.operator(
        "tesis.mark_landmark",
        text=_ui_landmark_label(idx),
        icon='RADIOBUT_ON' if state.active_landmark == idx else 'BLANK1',
    )
    op.lm_id = LANDMARKS_ESPERADOS[idx][0]

    value = split.row(align=True)
    value.alignment = 'RIGHT'
    value.label(text=_ui_landmark_value(lm))


def _draw_bilateral_pair_ui(layout, state, idx_l, idx_r, meshes_ok=True):
    """Dibuja un par bilateral en una fila: izquierda | derecha."""
    if idx_l >= len(state.landmarks) or idx_r >= len(state.landmarks):
        return

    row = layout.row(align=True)
    row.scale_y = 0.82
    if _landmark_ui_warn(state, idx_l) or _landmark_ui_warn(state, idx_r):
        row.alert = True

    split = row.split(factor=0.5, align=True)
    left = split.column(align=True)
    right = split.column(align=True)

    _draw_landmark_entry(left, state, idx_l, meshes_ok=meshes_ok)
    _draw_landmark_entry(right, state, idx_r, meshes_ok=meshes_ok)


def _draw_landmark_grid(layout, state, meshes_ok=True):
    """Tabla compacta de los 26 landmarks faciales.

    Centrales: dos columnas de seis.
    Bilaterales: siete filas, izquierda y derecha enfrentadas.
    No se muestran botones N/D repetidos dentro de cada casilla.
    """
    summary = f"{state.lm_count_set} medidos · {state.lm_count_missing} N/D"
    if state.lm_count_pending > 0:
        summary += f" · {state.lm_count_pending} pendientes"

    hdr = layout.row(align=True)
    hdr.scale_y = 0.82
    hdr.label(text=summary, icon='BONE_DATA')

    # CENTRALES ---------------------------------------------------
    title = layout.row(align=True)
    title.scale_y = 0.78
    title.label(text="CENTRALES")

    split = layout.split(factor=0.5, align=True)
    left_col = split.column(align=True)
    right_col = split.column(align=True)

    for offset in range(6):
        _draw_landmark_entry(
            left_col, state, _CENTRAL_INDICES[offset], meshes_ok=meshes_ok
        )
        _draw_landmark_entry(
            right_col, state, _CENTRAL_INDICES[offset + 6], meshes_ok=meshes_ok
        )

    layout.separator(factor=0.25)

    # BILATERALES -------------------------------------------------
    title = layout.row(align=True)
    title.scale_y = 0.78
    title.label(text="BILATERALES")

    headings = layout.split(factor=0.5, align=True)
    left_heading = headings.row(align=True)
    right_heading = headings.row(align=True)
    left_heading.alignment = 'CENTER'
    right_heading.alignment = 'CENTER'
    left_heading.label(text="IZQUIERDA")
    right_heading.label(text="DERECHA")

    for idx_l, idx_r in _iter_bilateral_pairs():
        _draw_bilateral_pair_ui(
            layout, state, idx_l, idx_r, meshes_ok=meshes_ok
        )

def _draw_neck_compact(layout, state, case):
    """Vista compacta cervical con Anterior y Posterior independientes."""
    hdr = layout.row(align=True)
    hdr.scale_y = 0.8
    hdr.label(text="Cuello (CCU)", icon='SORT_DESC')
    if not case.has_neck or not state.landmarks:
        return

    lm_ant = find_landmark_by_name(state.landmarks, "neck_ant")
    lm_post = find_landmark_by_name(state.landmarks, "neck_post")
    skull, soft = obtener_mallas_paciente(quiet=True)
    meshes_ok = skull is not None and soft is not None

    row = layout.row(align=True)
    row.enabled = meshes_ok

    ant = row.operator(
        "tesis.mark_neck_point",
        text="Anterior",
        icon='CHECKMARK' if lm_ant and lm_ant.status == "SET" else 'CHECKBOX_DEHLT',
    )
    ant.target = "neck_ant"

    post = row.operator(
        "tesis.mark_neck_point",
        text="Posterior",
        icon='CHECKMARK' if lm_post and lm_post.status == "SET" else 'CHECKBOX_DEHLT',
    )
    post.target = "neck_post"

    ready = bool(
        lm_ant and lm_post and
        lm_ant.status == "SET" and lm_post.status == "SET"
    )
    actions = layout.row(align=True)
    actions.enabled = ready and meshes_ok
    actions.operator("tesis.create_ccu_ring", text="Anillo", icon='MESH_CIRCLE')
    actions.operator("tesis.confirm_ccu", text="CCU", icon='CHECKMARK')

def _draw_case_columns(row, values, icon=None):
    """Dibuja una fila de tabla sin dependencias externas y compatible con Blender 5."""

    grid = row.grid_flow(
        row_major=True,
        columns=len(values),
        even_columns=True,
        even_rows=True,
        align=True,
    )

    for position, value in enumerate(values):
        cell = grid.row(align=True)

        if position == 0:
            cell.alignment = 'LEFT'
        else:
            cell.alignment = 'CENTER'

        # Blender 5 no acepta icon=None
        if position == 0 and icon:
            cell.label(text=str(value), icon=icon)
        else:
            cell.label(text=str(value))

class TESIS_UL_Cases(UIList):

    def filter_items(self, context, data, propname):
        state = context.scene.tesis_state
        cases = getattr(data, propname)

        flt_flags = []
        flt_neworder = []

        for i, c in enumerate(cases):
            ok = _case_matches_active_filters(state, c)
            flt_flags.append(self.bitflag_filter_item if ok else 0)

        return flt_flags, flt_neworder

    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        c = item

        if self.layout_type in {'DEFAULT', 'COMPACT'}:
            row = layout.row(align=True)
            row.scale_y = 0.90
            sexo_txt = "F" if c.sexo.strip().upper().startswith("F") else (
                "M" if c.sexo.strip().upper().startswith("M") else c.sexo
            )
            facial_done = c.summary_set + c.summary_missing
            state = context.scene.tesis_state
            patient_label = _safe_str(c.paciente)
            if state.measurement_mode == "ERROR" and int(getattr(c, "error_order", 0) or 0) > 0:
                patient_label = f"{int(c.error_order):02d} · {patient_label}"
            if state.measurement_mode == "ERROR":
                status_label = "✓" if _error_session_complete(c, state) else "Pend."
            else:
                status_label = "✓" if _case_is_complete(c) else "Pend."
            _draw_case_columns(row, [
                patient_label, _safe_str(c.edad) or "—", sexo_txt or "—",
                _safe_str(c.grupo) or "—", f"{facial_done}/{len(_OTHER_INDICES)}",
                status_label,
            ], icon='USER')

        elif self.layout_type == 'GRID':
            layout.alignment = 'CENTER'
            layout.label(text="", icon='USER')

# =====================================================
# OPERATORS
# =====================================================

class TESIS_OT_EnsureUnits(Operator):
    bl_idname = "tesis.ensure_units"
    bl_label = "Unidades en mm (Unit Scale 0.001)"
    bl_description = "Configura las unidades de la escena a milímetros (scale_length = 0.001)"

    def execute(self, context):
        ensure_scene_units_mm()
        self.report({'INFO'}, "Unidades configuradas: Metric / mm / scale 0.001")
        return {'FINISHED'}


class TESIS_OT_ToggleCaseNeck(Operator):
    bl_idname = "tesis.toggle_case_neck"
    bl_label = "Tiene cuello"
    bl_description = "Activa o desactiva explícitamente el estado del caso sin borrar silenciosamente datos cervicales"
    value: BoolProperty(default=False) # type: ignore

    def invoke(self, context, event):
        state = context.scene.tesis_state
        if not state.cases or state.active_case >= len(state.cases):
            return {'CANCELLED'}
        case = state.cases[state.active_case]
        if not self.value and case.has_neck:
            has_cervical_data = bool(
                any(
                    idx < len(state.landmarks) and state.landmarks[idx].status == "SET"
                    for idx in _NECK_INDICES
                )
                or _safe_float(getattr(case, "ccu_mm", 0.0), 0.0) > 0
            )
            if has_cervical_data:
                return context.window_manager.invoke_confirm(
                    self,
                    event,
                    message="Este caso contiene datos cervicales. ¿Confirmas que el caso no tiene cuello?",
                )
        return self.execute(context)

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases or state.active_case >= len(state.cases):
            return {'CANCELLED'}
        case = state.cases[state.active_case]
        case.has_neck = bool(self.value)
        # No destruir landmarks/CCU en memoria al desmarcar. El snapshot de
        # guardado normaliza neck_ant/neck_post y CCU solo cuando se escribe Excel.
        state.dirty = True
        self.report({'INFO'}, "Estado cervical actualizado; los datos se normalizan al guardar.")
        return {'FINISHED'}


def _error_stratum(case):
    sex = _safe_str(getattr(case, "sexo", "")).upper()
    sex_key = "F" if sex.startswith("F") else ("M" if sex.startswith("M") else "")
    age = _parse_age_years(getattr(case, "edad", ""))
    if age is None or not sex_key:
        return None
    band = "6_11" if 6 <= age <= 11 else ("12_18" if 12 <= age <= 18 else None)
    return f"{sex_key}_{band}" if band else None


def _error_case_is_eligible(case, include_neck=True):
    """Elegibilidad estricta para una cohorte limpia y reproducible."""
    if _case_requires_review(case):
        return False
    if _safe_str(getattr(case, "obs", "")).strip() not in {"", "0"}:
        return False
    skull_path = _safe_str(getattr(case, "ruta_hueso", "")).strip()
    soft_path = _safe_str(getattr(case, "ruta_blando", "")).strip()
    if not skull_path or not soft_path:
        return False
    try:
        skull_path = bpy.path.abspath(skull_path)
        soft_path = bpy.path.abspath(soft_path)
    except Exception:
        pass
    if not os.path.isfile(skull_path) or not os.path.isfile(soft_path):
        return False
    if _error_stratum(case) is None:
        return False
    if int(getattr(case, "summary_set", 0)) != len(_OTHER_INDICES):
        return False
    if int(getattr(case, "summary_missing", 0)) != 0 or int(getattr(case, "summary_pend", 0)) != 0:
        return False
    if include_neck:
        if not bool(getattr(case, "has_neck", False)):
            return False
        if _safe_str(getattr(case, "summary_neck_ant_status", "OMIT")).upper() != "SET":
            return False
        if _safe_str(getattr(case, "summary_neck_post_status", "OMIT")).upper() != "SET":
            return False
        if _safe_float(getattr(case, "baseline_ccu_mm", 0.0), 0.0) <= 0:
            return False
    return True


def _stratified_error_selection(cases, n, seed, include_neck=True):
    groups = {"F_6_11": [], "M_6_11": [], "F_12_18": [], "M_12_18": []}
    for case in cases:
        if not _error_case_is_eligible(case, include_neck=include_neck):
            continue
        key = _error_stratum(case)
        if key in groups:
            groups[key].append(case)

    total = sum(len(values) for values in groups.values())
    if total < n:
        raise ValueError(
            f"Solo hay {total} casos elegibles para una cohorte de {n}. "
            "Reduce N o desactiva 'Incluir CCU'."
        )

    # Asignación proporcional a la disponibilidad de casos elegibles.
    raw = {key: n * len(values) / total for key, values in groups.items()}
    quotas = {key: min(len(groups[key]), int(math.floor(raw[key]))) for key in groups}
    remaining = n - sum(quotas.values())
    order = sorted(groups, key=lambda key: (raw[key] - math.floor(raw[key]), len(groups[key])), reverse=True)
    while remaining > 0:
        changed = False
        for key in order:
            if quotas[key] < len(groups[key]):
                quotas[key] += 1
                remaining -= 1
                changed = True
                if remaining == 0:
                    break
        if not changed:
            break

    rng = random.Random(int(seed))
    selected = []
    for key in ("F_6_11", "M_6_11", "F_12_18", "M_12_18"):
        pool = list(groups[key])
        rng.shuffle(pool)
        selected.extend(pool[:quotas[key]])
    rng.shuffle(selected)
    return selected[:n], quotas, groups


class TESIS_OT_GenerateErrorCohort(Operator):
    bl_idname = "tesis.generate_error_cohort"
    bl_label = "Crear / usar cohorte de error"
    bl_description = (
        "Selecciona una sola vez una muestra estratificada y la guarda en el Excel. "
        "Si la cohorte ya existe, la reutiliza sin volver a sortear"
    )

    def execute(self, context):
        global _FILTER_SELECTION_SUPPRESS
        state = context.scene.tesis_state
        if not state.cases:
            self.report({'WARNING'}, "Carga primero el Excel.")
            return {'CANCELLED'}
        if state.dirty:
            self.report({'WARNING'}, "Guarda o descarta los cambios antes de crear la cohorte.")
            return {'CANCELLED'}

        cohort = _safe_str(state.error_cohort).strip() or _DEFAULT_ERROR_COHORT
        existing = [case for case in state.cases
                    if bool(getattr(case, "error_study", False))
                    and _safe_str(getattr(case, "error_cohort", "")) == cohort]
        if existing:
            invalid_existing = [
                case for case in existing
                if not _error_case_is_eligible(
                    case, include_neck=bool(state.error_include_neck)
                )
            ]
            if invalid_existing:
                self.report(
                    {'ERROR'},
                    f"La cohorte {cohort} contiene {len(invalid_existing)} casos que no cumplen "
                    "los criterios de cohorte limpia. Usa un identificador nuevo para sortearla."
                )
                return {'CANCELLED'}
            _FILTER_SELECTION_SUPPRESS = True
            try:
                state.error_cohort = cohort
                state.measurement_mode = "ERROR"
                state.error_only_pending = True
            finally:
                _FILTER_SELECTION_SUPPRESS = False
            _select_first_visible_case(context, sync=True)
            self.report({'INFO'}, f"Cohorte {cohort} reutilizada: {len(existing)} casos. No se volvió a sortear.")
            return {'FINISHED'}

        try:
            selected, quotas, groups = _stratified_error_selection(
                state.cases,
                int(state.error_sample_size),
                int(state.error_seed),
                include_neck=bool(state.error_include_neck),
            )
        except ValueError as exc:
            self.report({'ERROR'}, str(exc))
            return {'CANCELLED'}

        selected_ids = {_current_case_key(case): order for order, case in enumerate(selected, start=1)}
        excel_path = get_excel_path()
        _ensure_excel_columns(excel_path)
        with open_workbook(excel_path, save=True) as wb:
            ws = wb[SHEET_CASES]
            headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
            row_index = _build_case_row_index(ws, headers)
            for case in state.cases:
                row = _find_case_row_prefer_operational(
                    ws, case.paciente, case.ojo, headers, row_index
                )
                if row is None:
                    continue
                eligible = _error_case_is_eligible(case, include_neck=bool(state.error_include_neck))
                _ws_set_value(ws, row, headers[COL_ERROR_ELEGIBLE], 1 if eligible else 0)
                case.error_eligible = eligible
                key = _current_case_key(case)
                if key in selected_ids:
                    _ws_set_value(ws, row, headers[COL_ERROR_ESTUDIO], 1)
                    _ws_set_value(ws, row, headers[COL_ERROR_COHORTE], cohort)
                    _ws_set_value(ws, row, headers[COL_ERROR_ORDEN], selected_ids[key])
                    case.error_study = True
                    case.error_cohort = cohort
                    case.error_order = int(selected_ids[key])

        invalidate_excel_cache()
        _FILTER_SELECTION_SUPPRESS = True
        try:
            state.error_cohort = cohort
            state.measurement_mode = "ERROR"
            state.error_only_pending = True
        finally:
            _FILTER_SELECTION_SUPPRESS = False
        _select_first_visible_case(context, sync=True)
        quota_text = ", ".join(f"{key}={quotas[key]}" for key in ("F_6_11", "M_6_11", "F_12_18", "M_12_18"))
        self.report({'INFO'}, f"Cohorte {cohort} creada con {len(selected)} casos ({quota_text}).")
        return {'FINISHED'}


class TESIS_OT_ResetFilters(Operator):
    bl_idname = "tesis.reset_filters"
    bl_label = "Limpiar filtros"
    bl_description = "Restablece búsqueda, estado, sexo, edad, grupo, cuello, alerta y observaciones"

    def execute(self, context):
        state = context.scene.tesis_state
        global _FILTER_SELECTION_SUPPRESS
        _FILTER_SELECTION_SUPPRESS = True
        try:
            state.search = ""
            state.case_status = "PENDING"
            state.filter_sex = "ALL"
            state.filter_age_min = 0
            state.filter_age_max = 0
            state.filter_age_band = "ALL"
            state.filter_neck = "ALL"
            state.filter_eye = "ALL"
            state.filter_has_obs = "ALL"
            state.filter_has_lm_excel = False
        finally:
            _FILTER_SELECTION_SUPPRESS = False
        _select_first_visible_case(context, sync=True)
        self.report({'INFO'}, "Filtros restablecidos")
        return {'FINISHED'}


class TESIS_OT_LoadExcel(Operator):
    bl_idname = "tesis.load_excel"
    bl_label = "Cargar Casos desde Excel"
    bl_description = "Carga la lista de pacientes y casos desde el archivo Excel de datos"

    _result_cases = None
    _result_summaries = None
    _error = None

    def _apply_results(self, context):
        """Aplica los datos leídos al estado de Blender en el hilo principal."""
        state = context.scene.tesis_state
        _clear_case_visual_objects()
        _OBJ_NAME_TO_LM.clear()
        state.cases.clear()
        state.landmarks.clear()
        global _ACTIVE_CASE_CHANGE_SUPPRESS
        _ACTIVE_CASE_CHANGE_SUPPRESS = True
        try:
            state.active_case = 0
            state.committed_active_case = -1
        finally:
            _ACTIVE_CASE_CHANGE_SUPPRESS = False

        cases = TESIS_OT_LoadExcel._result_cases
        all_summaries = TESIS_OT_LoadExcel._result_summaries or {}

        if cases is None:
            err = TESIS_OT_LoadExcel._error
            self.report({'ERROR'}, str(err) if err else "Error desconocido al leer Excel.")
            return

        global _CASE_PROPERTY_UPDATE_SUPPRESS
        _CASE_PROPERTY_UPDATE_SUPPRESS = True
        try:
            for c in cases:
                item = state.cases.add()
                item.paciente = c["paciente"]
                item.ojo = c["ojo"]
                item.edad = c["edad"]
                item.sexo = c["sexo"]
                item.grupo = c["grupo"]
                item.peso = c["peso"]
                item.obs = c["obs"]
                item.ruta_blando = c["ruta_blando"]
                item.ruta_hueso = c["ruta_hueso"]
                item.procesado = c["procesado"]
                item.has_neck = bool(c.get("has_neck", False))
                item.faltan_landmarks = bool(c.get("faltan_landmarks", False))
                item.error_study = bool(c.get("error_study", False))
                item.error_cohort = _safe_str(c.get("error_cohort", ""))
                item.error_order = int(c.get("error_order", 0) or 0)
                item.error_eligible = bool(c.get("error_eligible", False))
                item.ccu_mm = 0.0
                item.estado_nutricional = ""
                item.baseline_ccu_mm = 0.0
                item.baseline_estado_nutricional = ""
                item.summary_neck_ant_status = "OMIT"
                item.summary_neck_post_status = "OMIT"

                summary = all_summaries.get(_case_pair_key(item.paciente, item.ojo))
                if summary:
                    item.summary_set = int(summary.get("set") or 0)
                    item.summary_missing = int(summary.get("omit") or 0)
                    item.summary_pend = int(summary.get("pend") or 0)
                    item.summary_n_valid = int(summary.get("n_valid") or 0)
                    item.summary_avg = float(summary.get("avg") or 0)
                    item.summary_median = float(summary.get("median") or 0)
                    item.summary_sd = float(summary.get("sd") or 0)
                    item.summary_min = float(summary.get("min") or 0)
                    item.summary_max = float(summary.get("max") or 0)
                    item.summary_loaded = True
                    item.summary_neck_ant_status = str(summary.get("neck_ant_status", "OMIT")).upper()
                    item.summary_neck_post_status = str(summary.get("neck_post_status", "OMIT")).upper()
                    ccu_summary = _safe_float(summary.get("ccu_mm"), 0.0)
                    if ccu_summary > 0:
                        item.ccu_mm = ccu_summary
                        item.baseline_ccu_mm = ccu_summary
                    if summary.get("estado_nutricional"):
                        item.estado_nutricional = str(summary["estado_nutricional"])
                        item.baseline_estado_nutricional = str(summary["estado_nutricional"])
        finally:
            _CASE_PROPERTY_UPDATE_SUPPRESS = False

        _update_case_counts(state)
        state.dirty = False
        if state.cases:
            error_cases = [
                case for case in state.cases
                if bool(getattr(case, "error_study", False))
                and _safe_str(getattr(case, "error_cohort", "")).strip()
            ]
            if error_cases:
                available_cohorts = sorted({
                    _safe_str(case.error_cohort).strip() for case in error_cases
                })
                # Elegir contexto ERROR por progreso REAL guardado, no por
                # orden alfabético. Esto evita abrir una cohorte vacía cuando
                # ya existen remediciones completas en otra cohorte.
                observer = _safe_str(state.observer_id).upper() or "OBS1"
                progress_by_context = {}
                for key, entry in _ERROR_PROGRESS_CACHE.items():
                    case_id, key_observer, key_session, key_replica, key_cohort = key
                    if key_observer != observer or key_cohort not in available_cohorts:
                        continue
                    statuses = entry.get("statuses", {}) if isinstance(entry, dict) else {}
                    facial_resolved = sum(
                        1 for idx in _OTHER_INDICES
                        if _safe_str(statuses.get(LANDMARKS_ESPERADOS[idx][0], "OMIT")).upper()
                        in {"SET", "MISSING"}
                    )
                    # Cada caso aporta primero por completitud facial y luego por
                    # cantidad resuelta. Así un contexto realmente trabajado
                    # gana sobre una cohorte recién creada pero aún vacía.
                    context_key = (key_cohort, _safe_str(key_session).upper(), int(key_replica or 1))
                    score = progress_by_context.setdefault(context_key, [0, 0, 0])
                    score[0] += 1 if facial_resolved == len(_OTHER_INDICES) else 0
                    score[1] += facial_resolved
                    score[2] += 1

                current_cohort = _safe_str(getattr(state, "error_cohort", "")).strip()
                current_session = _safe_str(getattr(state, "session_id", "S1")).upper()
                current_replica = int(getattr(state, "replicate_id", 1) or 1)
                current_context = (current_cohort, current_session, current_replica)

                if current_context in progress_by_context:
                    selected_context = current_context
                elif progress_by_context:
                    selected_context = max(
                        progress_by_context,
                        key=lambda ctx: (
                            progress_by_context[ctx][0],
                            progress_by_context[ctx][1],
                            progress_by_context[ctx][2],
                            ctx[0], ctx[1], -ctx[2],
                        ),
                    )
                else:
                    fallback_cohort = (
                        current_cohort if current_cohort in available_cohorts
                        else available_cohorts[0]
                    )
                    selected_context = (fallback_cohort, current_session or "S1", current_replica)

                selected_cohort, selected_session, selected_replica = selected_context
                global _FILTER_SELECTION_SUPPRESS
                _FILTER_SELECTION_SUPPRESS = True
                try:
                    state.error_cohort = selected_cohort
                    state.session_id = selected_session
                    state.replicate_id = selected_replica
                finally:
                    _FILTER_SELECTION_SUPPRESS = False

                if progress_by_context:
                    score = progress_by_context.get(selected_context, [0, 0, 0])
                    log.info(
                        "Contexto ERROR restaurado desde progreso: cohorte=%s observador=%s "
                        "sesion=%s replica=%s casos_completos=%d casos_con_datos=%d",
                        selected_cohort, observer, selected_session, selected_replica,
                        score[0], score[2],
                    )
            # committed=-1 fuerza la primera sincronización real del primer caso visible.
            _select_first_visible_case(context, sync=True)
        else:
            state.committed_active_case = -1
            _init_landmarks_for_case(state)
        state.dirty = False

        self.report({'INFO'}, f"{len(cases)} casos cargados desde Excel.")

    def execute(self, context):
        """Carga Excel de forma síncrona en el hilo principal de Blender.

        Todo el I/O de Excel se mantiene en el hilo principal de Blender. El patrón
        anterior de trabajo concurrente podía provocar cierres nativos mientras
        Blender seguía usando ``bpy`` o redibujando la UI. Para este libro (~300
        casos) se prioriza estabilidad y consistencia del estado.
        """
        state = context.scene.tesis_state

        ensure_scene_units_mm()

        if state.dirty and not bool(getattr(self, "_discard_confirmed", False)):
            self.report({'WARNING'}, "Hay cambios sin guardar. Confirma la recarga para descartarlos.")
            return {'CANCELLED'}

        excel_path = get_excel_path()
        self._excel_path = excel_path

        try:
            if (
                _EXCEL_CACHE["cases"] is not None
                and _EXCEL_CACHE["summaries"] is not None
                and not _excel_file_changed(excel_path)
            ):
                cases = _EXCEL_CACHE["cases"]
                summaries = _EXCEL_CACHE["summaries"]
            else:
                cases, summaries = load_cases_and_summaries(excel_path)

            TESIS_OT_LoadExcel._result_cases = cases
            TESIS_OT_LoadExcel._result_summaries = summaries
            TESIS_OT_LoadExcel._error = None
        except Exception as exc:
            TESIS_OT_LoadExcel._result_cases = None
            TESIS_OT_LoadExcel._result_summaries = None
            TESIS_OT_LoadExcel._error = exc
            log.exception("Error cargando Excel ruta=%s", excel_path)

        self._apply_results(context)
        try:
            for area in context.screen.areas:
                area.tag_redraw()
        except Exception:
            pass

        return {'FINISHED'} if TESIS_OT_LoadExcel._error is None else {'CANCELLED'}

    def invoke(self, context, event):
        state = context.scene.tesis_state
        if state.dirty:
            self._discard_confirmed = True
            return context.window_manager.invoke_confirm(
                self, event,
                message="Hay cambios sin guardar. ¿Recargar Excel y descartarlos?"
            )
        self._discard_confirmed = False
        return self.execute(context)

    def modal(self, context, event):
        # Compatibilidad defensiva: la carga ya no usa modal ni threads.
        return {'FINISHED'}


def _init_landmarks_for_case(state: TesisState):
    state.landmarks.clear()

    for lm_id, _soft_id, lm_label in LANDMARKS_ESPERADOS:
        lm = state.landmarks.add()
        lm.name = lm_id
        lm.status = "OMIT"
        lm.x = 0.0
        lm.y = 0.0
        lm.z = 0.0
        lm.hx = 0.0
        lm.hy = 0.0
        lm.hz = 0.0
        lm.thickness = 0.0
        lm.hard_obj_name = ""
        lm.soft_obj_name = ""
        lm.line_obj_name = ""

    state.active_landmark = 0
    _update_lm_counts(state)


def _apply_legacy_import_orientation(obj):
    """Hornea la orientación histórica del addon: X +90° y luego Y -90°.

    Solo se aplica cuando la persona la selecciona explícitamente. No se usa por
    defecto para evitar rotar dos veces PLY que ya fueron exportados orientados.
    """
    if obj is None or getattr(obj, "type", None) != 'MESH':
        return False
    try:
        rot_x = Matrix.Rotation(math.radians(90.0), 4, 'X')
        rot_y = Matrix.Rotation(math.radians(-90.0), 4, 'Y')
        # Equivale a transformar primero por X y después por Y.
        obj.data.transform(rot_y @ rot_x)
        obj.data.update()
        obj.rotation_euler = (0.0, 0.0, 0.0)
        obj.scale = (1.0, 1.0, 1.0)
        obj["tesis_import_orientation"] = "LEGACY_X90_YM90"
        return True
    except Exception as exc:
        log.warning("No se pudo aplicar orientación legado a %s: %s", getattr(obj, "name", "?"), exc)
        return False


# =============================================================================
# IMPORTACIÓN DE MALLAS DEL CASO ACTIVO
# =============================================================================
# El caso activo se importa a Blender con dos mallas principales:
# - SKULL: superficie ósea;
# - SOFT: superficie de tejido blando.
#
# Después de la carga, el programa valida si las coordenadas históricas del caso
# son compatibles con la geometría actual y reconstruye los landmarks ya guardados.
#
# Este bloque encapsula la etapa de preparación del caso para la medición.
# =============================================================================
class TESIS_OT_ImportMeshes(Operator):
    _busy = False
    _timer = None

    bl_idname = "tesis.import_meshes"
    bl_label = "Importar Mallas del Caso"
    bl_description = "Importa las mallas de hueso y tejido blando del caso activo"

    def execute(self, context):
        if type(self)._busy:
            self.report({'WARNING'}, "Importación ya en curso")
            return {'CANCELLED'}
        state = context.scene.tesis_state

        if not state.cases:
            self.report({'WARNING'}, "No hay casos cargados.")
            return {'CANCELLED'}

        if context.mode != 'OBJECT':
            try:
                bpy.ops.object.mode_set(mode='OBJECT')
            except RuntimeError:
                pass

        case = state.cases[state.active_case]
        for label, path in [("Hueso", case.ruta_hueso), ("Blando", case.ruta_blando)]:
            if not path or not path.strip():
                self.report({'ERROR'}, f"Ruta de {label} vacía para {case.paciente}/{case.ojo}.")
                return {'CANCELLED'}
            if not os.path.exists(path):
                self.report({'ERROR'}, f"Archivo {label} no encontrado: {path}")
                return {'CANCELLED'}

        type(self)._busy = True
        context.window.cursor_set('WAIT')
        wm = context.window_manager
        wm.progress_begin(0, 100)
        perf = {}
        started = time.perf_counter()
        created_objects = []
        skull = None
        soft = None
        landmarks_count = 0

        try:
            wm.progress_update(5)
            with perf_section("limpiar caso", perf):
                _clear_case_visual_objects()
            _OBJ_NAME_TO_LM.clear()
            invalidate_bvh_cache()

            wm.progress_update(20)
            with perf_section("SKULL import", perf):
                skull = _import_ply_and_get_object(context, case.ruta_hueso)
            if skull is None:
                raise RuntimeError("No pude identificar la malla de hueso importada.")
            created_objects.append(skull)
            skull.name = f"SKULL_{case.paciente}_{case.ojo}"
            _tag_case_mesh(skull, "SKULL", case)

            wm.progress_update(45)
            with perf_section("SOFT import", perf):
                soft = _import_ply_and_get_object(context, case.ruta_blando)
            if soft is None:
                raise RuntimeError("No pude identificar la malla de tejido blando importada.")
            created_objects.append(soft)
            soft.name = f"SOFT_{case.paciente}_{case.ojo}"
            _tag_case_mesh(soft, "SOFT", case)

            with perf_section("leer Excel", perf):
                # Las coordenadas canónicas se usan únicamente para validar el marco
                # geométrico. En modo ERROR no se reconstruyen para el observador.
                canonical_data = load_landmarks_from_excel(case)
                measurement_data = _load_active_measurement_data(case, state)
            with perf_section("construir BVH SKULL/SOFT", perf):
                ensure_bvh_cache(context, skull, soft, perf=perf)
            with perf_section("validar coordenadas históricas", perf):
                orientation, errors, frame_valid, validation_message = validate_saved_landmarks_against_meshes(
                    canonical_data, skull, soft, cache_key=_frame_validation_key(case, skull, soft),
                    patient=case.paciente,
                )
            if not frame_valid or orientation is None:
                case.coordinate_frame_valid = False
                case.coordinate_review_required = True
                cols = ensure_tesis_collections()
                move_object_to_collection(skull, cols["meshes"])
                move_object_to_collection(soft, cols["meshes"])
                _ensure_object_visible(skull)
                _ensure_object_visible(soft)
                cols["root"].hide_viewport = False
                cols["meshes"].hide_viewport = False
                self.report({'WARNING'}, "Coordenadas históricas no alineadas. Mallas cargadas solo para revisión.")
                landmarks_count = 0
                wm.progress_update(100)
                return {'FINISHED'}

            if orientation == "LEGACY":
                _apply_legacy_import_orientation(skull)
                _apply_legacy_import_orientation(soft)
                invalidate_bvh_cache()

            case.coordinate_frame = _COORDINATE_FRAME
            case.import_orientation = orientation
            case.coordinate_frame_valid = True
            case.coordinate_review_required = False
            state.import_orientation = orientation
            skull["coordinate_frame"] = _COORDINATE_FRAME
            soft["coordinate_frame"] = _COORDINATE_FRAME
            skull["tesis_import_orientation"] = orientation
            soft["tesis_import_orientation"] = orientation
            if "revisar" in validation_message.lower():
                self.report({'WARNING'}, validation_message)

            cols = ensure_tesis_collections()
            move_object_to_collection(skull, cols["meshes"])
            move_object_to_collection(soft, cols["meshes"])
            _ensure_object_visible(skull)
            _ensure_object_visible(soft)
            try:
                cols["root"].hide_viewport = False
                cols["meshes"].hide_viewport = False
            except Exception:
                pass
            try:
                bpy.ops.object.select_all(action='DESELECT')
                context.view_layer.objects.active = None
            except Exception:
                pass

            with perf_section("BVH preparación", perf):
                ensure_bvh_cache(context, skull, soft, perf=perf)
            wm.progress_update(65)
            wm.progress_update(85)
            with perf_section("LM rebuild", perf):
                landmarks_count = rebuild_landmarks_from_data(context, case, state, data=measurement_data)
            if case.reconstruction_warning:
                self.report({'WARNING'}, case.reconstruction_warning)

            if getattr(case, "ccu_mm", 0.0) > 0:
                log.info(
                    "CCU histórica conservada para %s; anillo visual diferido hasta usar la sección CCU.",
                    case.paciente,
                )

            wm.progress_update(100)
        except Exception as e:
            log.exception(
                "Error importando paciente=%s fase=importación skull=%s soft=%s",
                case.paciente, case.ruta_hueso, case.ruta_blando,
            )
            for obj in created_objects:
                _safe_remove_object(obj)
            invalidate_bvh_cache()
            self.report({'ERROR'}, f"Error al importar mallas: {e}")
            return {'CANCELLED'}
        finally:
            try:
                bpy.ops.object.select_all(action='DESELECT')
                context.view_layer.objects.active = None
            except Exception:
                pass
            wm.progress_end()
            context.window.cursor_set('DEFAULT')
            perf["TOTAL"] = time.perf_counter() - started
            _print_perf_block("IMPORT", perf)
            _memory_diagnostic(_current_case_key(case))
            type(self)._busy = False

        msg = "Mallas importadas"
        if landmarks_count > 0:
            msg += f" | {landmarks_count} landmarks reconstruidos desde Excel"
        else:
            msg += " | sin landmarks previos guardados"
        self.report({'INFO'}, msg)
        return {'FINISHED'}


class TESIS_OT_LoadSavedLandmarks(Operator):

    bl_idname = "tesis.load_saved_landmarks"
    bl_label = "Reconstruir objetos LM desde Excel"
    bl_description = "Vuelve a crear las esferas y líneas del caso desde las coordenadas ya guardadas. No modifica el Excel."

    def execute(self, context):
        state = context.scene.tesis_state

        if not state.cases:
            self.report({'WARNING'}, "No hay casos cargados.")
            return {'CANCELLED'}

        case = state.cases[state.active_case]
        skull, soft = obtener_mallas_paciente(quiet=True)
        if skull is None or soft is None:
            self.report({'WARNING'}, "Importa primero las mallas del caso activo.")
            return {'CANCELLED'}

        landmarks_count = rebuild_landmarks_from_data(context, case, state)
        if landmarks_count > 0:
            self.report({'INFO'}, f"{landmarks_count} landmarks reconstruidos desde Excel.")
        else:
            self.report({'INFO'}, "No hay landmarks guardados para este caso.")
        return {'FINISHED'}

# Flag módulo: True mientras el modal de marcado secuencial está activo
_marking_modal_active = False


def _clear_view3d_headers():
    """Limpia el texto de cabecera en todas las áreas VIEW_3D visibles."""
    try:
        screen = bpy.context.screen
        if screen is None:
            return
        for area in screen.areas:
            if area.type == 'VIEW_3D':
                area.header_text_set(None)
    except Exception:
        pass


def _update_marking_header_from_state(state):
    """Actualiza la cabecera del viewport con el landmark activo y atajos."""
    if state is None or not getattr(state, "landmarks", None):
        return
    idx = int(getattr(state, "active_landmark", -1))
    if idx < 0 or idx >= len(LANDMARKS_ESPERADOS):
        return
    lm_id = LANDMARKS_ESPERADOS[idx][0]
    text = (
        f"▶ {idx + 1}.{lm_id} | Click = marcar | "
        "N = no disponible | ESC = pausar | ◀▶ navegar"
    )
    try:
        screen = bpy.context.screen
        if screen is None:
            return
        for area in screen.areas:
            if area.type == 'VIEW_3D':
                area.header_text_set(text)
    except Exception:
        pass


def _finish_marking_modal():
    """Cierra el estado global del modal de marcado y limpia cabeceras."""
    global _marking_modal_active
    _marking_modal_active = False
    _clear_view3d_headers()

def _reset_landmark_to_pending_runtime(state, idx: int) -> bool:
    """Deshace un landmark dentro del modal sin usar el Undo global de Blender."""
    if idx < 0 or idx >= len(state.landmarks):
        return False
    lm = state.landmarks[idx]

    # Desvincular únicamente los objetos de este landmark.
    names = {
        _safe_str(getattr(lm, "hard_obj_name", "")),
        _safe_str(getattr(lm, "soft_obj_name", "")),
        _safe_str(getattr(lm, "line_obj_name", "")),
    }
    for name in names:
        if not name:
            continue
        obj = bpy.data.objects.get(name)
        if obj is not None:
            _safe_remove_object(obj)

    # Fallback por prefijo, por si algún nombre no quedó registrado.
    prefix = f"{idx:02d}_"
    for obj in list(bpy.data.objects):
        if obj.name.startswith(prefix):
            _safe_remove_object(obj)

    _unregister_lm_objs(idx)
    lm.status = "OMIT"
    lm.hx = lm.hy = lm.hz = 0.0
    lm.x = lm.y = lm.z = 0.0
    lm.thickness = 0.0
    lm.hard_obj_name = ""
    lm.soft_obj_name = ""
    lm.line_obj_name = ""
    state.active_landmark = idx
    state.dirty = True
    _update_lm_counts(state)
    return True


# =============================================================================
# MARCADO DE LANDMARKS
# =============================================================================
# El flujo de marcado sigue el principio de separación entre:
# - decisión anatómica del observador;
# - automatización geométrica para localizar la superficie de tejido blando;
# - comprobación visual del grosor calculado.
#
# En esencia, el usuario reconoce el punto en la estructura ósea y el sistema
# busca la proyección válida hacia el tejido blando, generando la distancia del
# espesor. Si la superficie no es compatible o la medición es dudosa, el usuario
# puede corregirla manualmente.
# =============================================================================
class TESIS_OT_AddLandmark(Operator):
    bl_idname = "tesis.add_landmark"
    bl_label = "Marcar Landmark (click cráneo -> blando)"
    bl_description = "Modo interactivo: click en malla de hueso para posicionar landmark y calcular thickness hasta tejido blando"
    # Undo global no es seguro mientras este modal crea/elimina varios objetos.
    # Ctrl+Z se maneja internamente como 'deshacer último landmark'.
    bl_options = {'REGISTER'}

    sequential: BoolProperty(default=True) # type: ignore
    start_from_active: BoolProperty(default=False) # type: ignore

    @classmethod
    def poll(cls, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return False
        skull, soft = obtener_mallas_paciente(quiet=True)
        return skull is not None and soft is not None

    def invoke(self, context, event):
        state = context.scene.tesis_state

        if not state.cases:
            self.report({'WARNING'}, "No hay caso activo.")
            return {'CANCELLED'}

        if not hasattr(state, "landmarks") or not state.landmarks:
            self.report({'ERROR'}, "No hay lista de landmarks cargada.")
            return {'CANCELLED'}

        skull, soft = obtener_mallas_paciente()
        if not skull or not soft:
            self.report({'ERROR'}, "Importa mallas primero (SKULL_ y SOFT_) o asigna en panel.")
            return {'CANCELLED'}

        try:
            ensure_bvh_cache(context, skull, soft)
        except Exception as e:
            self.report({'ERROR'}, f"Error construyendo BVH: {e}")
            return {'CANCELLED'}

        if state.active_landmark < 0 or state.active_landmark >= len(state.landmarks):
            state.active_landmark = 0

        global _marking_modal_active
        _marking_modal_active = True
        self._paused = False
        self._mark_history = []

        # El inicio global busca el primer pendiente; el clic directo conserva el índice.
        if self.sequential:
            if not self.start_from_active:
                state.active_landmark = 0
                # Buscar primer pendiente en orden secuencial
                found_pending = False
                for si in _SEQUENTIAL_ORDER:
                    if state.landmarks[si].status == "OMIT":
                        state.active_landmark = si
                        found_pending = True
                        break
                if not found_pending:
                    _finish_marking_modal()
                    self.report({'INFO'}, "\u2705 Todos los landmarks ya están marcados.")
                    return {'CANCELLED'}

        context.window_manager.modal_handler_add(self)
        idx = state.active_landmark
        lm_id = LANDMARKS_ESPERADOS[idx][0]
        if context.area:
            context.area.header_text_set(
                f"\u25B6 {idx+1}.{lm_id} | Click = marcar | N = no disponible | ESC = pausar | \u25C0\u25B6 navegar"
            )
        return {'RUNNING_MODAL'}

    def modal(self, context, event):
        global _marking_modal_active

        state = context.scene.tesis_state
        if not state.cases or not state.landmarks:
            _finish_marking_modal()
            return {'CANCELLED'}
        if state.active_landmark < 0 or state.active_landmark >= len(state.landmarks):
            state.active_landmark = 0

        # Ctrl+Z durante el modal NO entra al Undo global de Blender.
        # Deshace solo el último landmark creado en esta sesión, evitando invalidar
        # objetos/materiales mientras el operador modal todavía está vivo.
        if event.type == 'Z' and event.ctrl and event.value == 'PRESS':
            history = getattr(self, '_mark_history', [])
            if not history:
                self.report({'INFO'}, "No hay landmarks nuevos de esta sesión para deshacer.")
                return {'RUNNING_MODAL'}
            idx_undo = history.pop()
            if _reset_landmark_to_pending_runtime(state, idx_undo):
                lm_id_undo = LANDMARKS_ESPERADOS[idx_undo][0]
                msg = f"↶ Deshecho {idx_undo + 1}.{lm_id_undo} — vuelve a marcarlo"
                if context.area:
                    context.area.header_text_set(msg)
                self.report({'INFO'}, msg)
            return {'RUNNING_MODAL'}

        # ── Estado pausado: solo SPACE reanuda, todo lo demás pasa a Blender ──
        if self._paused:
            if event.type == 'SPACE' and event.value == 'PRESS':
                self._paused = False
                _marking_modal_active = True
                idx = state.active_landmark
                lm_id = LANDMARKS_ESPERADOS[idx][0]
                if context.area:
                    context.area.header_text_set(
                        f"\u25B6 {idx+1}.{lm_id} | Click = marcar | N = no disponible | ESC = pausar | \u25C0\u25B6 navegar"
                    )
                self.report({'INFO'}, "\u25B6 Reanudado — click izquierdo para marcar.")
                return {'RUNNING_MODAL'}
            if event.type in {'ESC', 'RIGHTMOUSE'} and event.value == 'PRESS':
                _finish_marking_modal()
                self.report({'INFO'}, "\u23F8 Marcado finalizado.")
                return {'CANCELLED'}
            # Mientras está en pausa, dejar que Blender maneje todo lo demás
            return {'PASS_THROUGH'}

        # Atajo N = marcar el landmark facial activo como no disponible.
        # Solo se intercepta mientras el marcado está ACTIVO; fuera del modal,
        # N conserva su función normal de Blender (mostrar/ocultar el panel lateral).
        if (
            event.type == 'N'
            and event.value == 'PRESS'
            and not event.ctrl and not event.alt and not event.shift
        ):
            idx = int(state.active_landmark)
            if idx not in _OTHER_INDICES:
                self.report({'WARNING'}, "El cuello se gestiona desde su sección propia.")
                return {'RUNNING_MODAL'}

            lm = state.landmarks[idx]

            # Para evitar borrar una medición existente por accidente, el hotkey N
            # solo actúa directamente sobre PENDIENTE o MISSING.
            if lm.status == "SET":
                self.report({
                    'WARNING'
                }, f"{_ui_landmark_label(idx)} ya está medido. Usa el botón del panel para confirmar que quieres marcarlo N/D.")
                return {'RUNNING_MODAL'}

            if lm.status == "MISSING":
                # N sobre un N/D lo reactiva como pendiente y se queda en ese LM.
                lm.status = "OMIT"
                lm.hx = lm.hy = lm.hz = 0.0
                lm.x = lm.y = lm.z = 0.0
                lm.thickness = 0.0
                lm.hard_obj_name = ""
                lm.soft_obj_name = ""
                lm.line_obj_name = ""
                state.dirty = True
                _update_lm_counts(state)
                _update_marking_header_from_state(state)
                self.report({'INFO'}, f"{_ui_landmark_label(idx)} volvió a pendiente.")
                return {'RUNNING_MODAL'}

            # OMIT -> MISSING
            clear_landmark_objects(idx)
            _unregister_lm_objs(idx)
            lm.status = "MISSING"
            lm.hx = lm.hy = lm.hz = 0.0
            lm.x = lm.y = lm.z = 0.0
            lm.thickness = 0.0
            lm.hard_obj_name = ""
            lm.soft_obj_name = ""
            lm.line_obj_name = ""
            state.dirty = True
            _update_lm_counts(state)

            # Permitir Ctrl+Z seguro dentro del modal para volverlo a pendiente.
            if hasattr(self, "_mark_history"):
                self._mark_history.append(idx)

            if self.sequential:
                next_idx = _next_sequential_idx(idx)
                if next_idx is not None:
                    state.active_landmark = next_idx

            pending = sum(
                1 for si in _SEQUENTIAL_ORDER
                if state.landmarks[si].status == "OMIT"
            )
            if pending == 0:
                _finish_marking_modal()
                self.report({'INFO'}, "✅ Landmarks faciales resueltos.")
                return {'FINISHED'}

            _update_marking_header_from_state(state)
            self.report({'INFO'}, f"{_ui_landmark_label(idx)} marcado N/D.")
            return {'RUNNING_MODAL'}

        # Pausar (ESC o RMB) — detiene sin perder progreso, SPACE reanuda
        if event.type in {'ESC', 'RIGHTMOUSE'} and event.value == 'PRESS':
            self._paused = True
            _marking_modal_active = False
            try:
                if context.area:
                    context.area.header_text_set("\u23F8 Pausado — SPACE para reanudar")
            except Exception:
                pass
            self.report({'INFO'}, "\u23F8 Pausado \u2014 navega la escena. SPACE para reanudar.")
            return {'RUNNING_MODAL'}

        # Click izquierdo: marcar landmark (PRESS o CLICK)
        if event.type == 'LEFTMOUSE' and event.value in {'PRESS', 'CLICK'}:
            # Verificar que estamos en viewport 3D
            if context.area and context.area.type == 'VIEW_3D':
                # Buscar la región WINDOW del viewport
                region = None
                for r in context.area.regions:
                    if r.type == 'WINDOW':
                        region = r
                        break
                if region and (region.x <= event.mouse_x <= region.x + region.width
                               and region.y <= event.mouse_y <= region.y + region.height):
                    marked_idx = state.active_landmark
                    marked_was_pending = (state.landmarks[marked_idx].status == "OMIT")
                    try:
                        ok = self._handle_click(context, event, region)
                    except Exception as exc:
                        log.error("Error en _handle_click: %s", exc, exc_info=True)
                        self.report({'ERROR'}, f"Error al marcar: {exc}")
                        return {'RUNNING_MODAL'}
                    if ok:
                        if marked_was_pending:
                            self._mark_history.append(marked_idx)
                        if not self.sequential:
                            _finish_marking_modal()
                            self.report({'INFO'}, "Landmark marcado.")
                            return {'FINISHED'}
                        # Contar pendientes solo de landmarks en flujo secuencial (sin cuello)
                        pending = sum(1 for si in _SEQUENTIAL_ORDER if state.landmarks[si].status == "OMIT")
                        if pending == 0:
                            _finish_marking_modal()
                            self.report({'INFO'}, "\u2705 Landmarks secuenciales completados! (marca cuello manualmente)")
                            return {'FINISHED'}
                        lm_id = LANDMARKS_ESPERADOS[state.active_landmark][0]
                        msg = f"\u25B6 {state.active_landmark + 1}.{lm_id} | Click = marcar | N = no disponible | ESC = pausar | \u25C0\u25B6 navegar"
                        if context.area:
                            context.area.header_text_set(msg)
                        self.report({'INFO'}, msg)
                    return {'RUNNING_MODAL'}

            # Click fuera del viewport — dejar pasar para interactuar con paneles
            return {'PASS_THROUGH'}

        # LEFTMOUSE RELEASE:
        # dentro del viewport se consume para evitar una segunda acción;
        # fuera de la región WINDOW (por ejemplo en el N-panel) se deja pasar,
        # de modo que los botones de la interfaz funcionen SIN pausar con ESC.
        if event.type == 'LEFTMOUSE' and event.value == 'RELEASE':
            if context.area and context.area.type == 'VIEW_3D':
                window_region = next(
                    (r for r in context.area.regions if r.type == 'WINDOW'),
                    None
                )
                if window_region is not None:
                    inside_window = (
                        window_region.x <= event.mouse_x <= window_region.x + window_region.width
                        and window_region.y <= event.mouse_y <= window_region.y + window_region.height
                    )
                    if inside_window:
                        return {'RUNNING_MODAL'}
            return {'PASS_THROUGH'}

        # Flechas izq/derecha: navegar entre landmarks sin salir del modal
        if event.type == 'LEFT_ARROW' and event.value == 'PRESS':
            state = context.scene.tesis_state
            prev_idx = _prev_sequential_idx(state.active_landmark)
            if prev_idx is not None:
                state.active_landmark = prev_idx
                lm_id = LANDMARKS_ESPERADOS[prev_idx][0]
                if context.area:
                    context.area.header_text_set(
                        f"\u25B6 {prev_idx+1}.{lm_id} | Click = marcar | N = no disponible | ESC = pausar | \u25C0\u25B6 navegar"
                    )
            return {'RUNNING_MODAL'}

        if event.type == 'RIGHT_ARROW' and event.value == 'PRESS':
            state = context.scene.tesis_state
            next_idx = _next_sequential_idx(state.active_landmark)
            if next_idx is not None:
                state.active_landmark = next_idx
                lm_id = LANDMARKS_ESPERADOS[next_idx][0]
                if context.area:
                    context.area.header_text_set(
                        f"\u25B6 {next_idx+1}.{lm_id} | Click = marcar | N = no disponible | ESC = pausar | \u25C0\u25B6 navegar"
                    )
            return {'RUNNING_MODAL'}

        # G debe poder iniciar el operador de movimiento aunque el marcado
        # secuencial siga vivo. Pausamos este modal antes de dejar pasar G para
        # que el LEFTMOUSE que confirma el movimiento NO marque otro landmark.
        if event.type == 'G' and event.value == 'PRESS':
            self._paused = True
            _marking_modal_active = False
            try:
                if context.area:
                    context.area.header_text_set(
                        "⏸ Marcado pausado para mover LM · SPACE para reanudar"
                    )
            except Exception:
                pass
            return {'PASS_THROUGH'}

        # X / DEL también deben llegar al keymap del borrado lógico.
        if event.type in {'X', 'DEL'} and event.value == 'PRESS':
            return {'PASS_THROUGH'}

        # Dejar pasar navegación viewport, scroll, y atajos de Blender (Ctrl+Z, etc.)
        if event.type in {'MIDDLEMOUSE', 'WHEELUPMOUSE', 'WHEELDOWNMOUSE',
                          'MOUSEMOVE', 'INBETWEEN_MOUSEMOVE',
                          'TRACKPADPAN', 'TRACKPADZOOM',
                          'NUMPAD_0', 'NUMPAD_1', 'NUMPAD_2', 'NUMPAD_3',
                          'NUMPAD_4', 'NUMPAD_5', 'NUMPAD_6', 'NUMPAD_7',
                          'NUMPAD_8', 'NUMPAD_9', 'NUMPAD_PERIOD',
                          'HOME', 'END'}:
            return {'PASS_THROUGH'}

        # Otros atajos Ctrl/Alt siguen disponibles. Ctrl+Z ya fue interceptado
        # arriba para usar el deshacer seguro del addon.
        if event.ctrl or event.alt:
            return {'PASS_THROUGH'}

        return {'RUNNING_MODAL'}

    def cancel(self, context):
        self._paused = False
        _finish_marking_modal()

    def _handle_click(self, context, event, viewport_region=None):
        perf = {}
        landmark_started = time.perf_counter()
        state = context.scene.tesis_state
        skull, soft = obtener_mallas_paciente()
        if not skull or not soft:
            self.report({'ERROR'}, "Importa mallas primero o asigna en panel.")
            return False

        with perf_section("BVH cache", perf):
            ensure_bvh_cache(context, skull, soft)
        bvh_skull = _BVH_CACHE["skull"]
        bvh_soft  = get_full_soft_bvh()
        if bvh_skull is None or bvh_soft is None:
            self.report({'ERROR'}, "BVH cache no disponible.")
            return False

        # Usar la región WINDOW del viewport explícitamente
        region = viewport_region
        if region is None:
            # Fallback: buscar región WINDOW
            if context.area and context.area.type == 'VIEW_3D':
                for r in context.area.regions:
                    if r.type == 'WINDOW':
                        region = r
                        break
        if region is None:
            self.report({'ERROR'}, "No se encontró región 3D.")
            return False

        rv3d = context.space_data.region_3d if context.space_data else None
        if rv3d is None:
            self.report({'ERROR'}, "Esto debe ejecutarse en Viewport 3D.")
            return False

        # Calcular coordenadas relativas a la región WINDOW desde coordenadas absolutas
        mouse_x = event.mouse_x - region.x
        mouse_y = event.mouse_y - region.y
        coord = (mouse_x, mouse_y)
        ray_origin = view3d_utils.region_2d_to_origin_3d(region, rv3d, coord)
        ray_dir = view3d_utils.region_2d_to_vector_3d(region, rv3d, coord).normalized()

        idx = state.active_landmark
        lm_id, _soft_id, lm_label = LANDMARKS_ESPERADOS[idx]
        lm_normalized = normalize_landmark_name(lm_id)
        
        # Detectar si es landmark de cuello
        is_neck_landmark = lm_normalized in _NECK_IDS
        
        if is_neck_landmark:
            # LANDMARKS DE CUELLO: solo marcar en tejido blando
            hit_loc, hit_norm, face_index, dist = bvh_soft.ray_cast(ray_origin, ray_dir)
            if hit_loc is None:
                self.report({'WARNING'}, "No intersecté el tejido blando.")
                return False
            
            p_soft = hit_loc
            p_hard = None  # No hay punto hard en landmarks de cuello
            
            # Guardar solo punto blando
            lm = state.landmarks[state.active_landmark]
            lm.status = "SET"
            lm.x, lm.y, lm.z = p_soft.x, p_soft.y, p_soft.z
            lm.hx, lm.hy, lm.hz = 0.0, 0.0, 0.0
            lm.thickness = 0.0
            
        else:
            # LANDMARKS NORMALES: punto hard + soft + thickness
            # 1) HIT EN CRÁNEO (BVH EN WORLD → rayo EN WORLD)
            with perf_section("SKULL raycast", perf):
                hit_loc, hit_norm, face_index, dist = bvh_skull.ray_cast(ray_origin, ray_dir)
            if hit_loc is None:
                self.report({'WARNING'}, "No intersecté el CRÁNEO con el click.")
                return False

            p_hard = hit_loc
            n_hard = hit_norm.normalized() if hit_norm is not None else ray_dir

            # 2) Normal outward (flip si apunta hacia el centro)
            n_out = outward_normal_at_point(skull, p_hard, n_hard)

            # Candidatos en la dirección ósea: solo se acepta el primero visible
            # desde el rayo original del viewport.
            with perf_section("SOFT projection", perf):
                candidates = ray_cast_all_hits(
                    bvh_soft, p_hard + n_out * 0.05, n_out,
                    max_dist=500.0, epsilon=0.05,
                )
            with perf_section("visibility check", perf):
                valid_candidates = [
                    hit for hit in candidates
                    if hit["distance"] > 0.05
                    and is_soft_candidate_visible_from_view(bvh_soft, ray_origin, hit["point"])
                ]
            if not valid_candidates:
                self.report({'WARNING'}, f"No se encontró piel externa válida para {lm_id}. Posicione manualmente.")
                return False
            p_soft = min(valid_candidates, key=lambda hit: hit["distance"])["point"]

            # 4) Guardar en LandmarkItem
            lm = state.landmarks[state.active_landmark]
            lm.status = "SET"

            lm.hx, lm.hy, lm.hz = p_hard.x, p_hard.y, p_hard.z
            lm.x,  lm.y,  lm.z  = p_soft.x, p_soft.y, p_soft.z
            lm.thickness = (p_soft - p_hard).length
            
            # Validación científica de thickness
            if lm.thickness < 1.0:
                self.report({'WARNING'}, f"{lm_id}: thickness = {lm.thickness:.2f} mm sospechosamente bajo. Verificar geometría.")
            elif lm.thickness > 50.0:
                self.report({'WARNING'}, f"{lm_id}: thickness = {lm.thickness:.2f} mm anómalamente alto. ¿Capturó artefacto?")

        nice_name = f"{idx:02d}_{lm_id}"

        # Limpiar todos los landmarks previos de este índice
        objects_started = time.perf_counter()
        clear_landmark_objects(idx)
        
        cols = ensure_tesis_collections()
        
        if is_neck_landmark:
            # LANDMARKS DE CUELLO: solo esfera en tejido blando (verde)
            soft_obj = create_landmark_sphere(
                f"{nice_name}_SOFT",
                p_soft,
                1.5,  # Un poco más grande para distinguir
                (0, 1, 0.3, 1),  # Verde para landmarks de cuello
                cols["soft"].name
            )
            soft_obj.hide_select = False
            soft_obj.lock_location = (False, False, False)
            _remove_soft_surface_constraint(soft_obj)
            
            lm.soft_obj_name = soft_obj.name
            lm.hard_obj_name = ""  # No hay objeto hard
            lm.line_obj_name = ""  # No hay línea
            _register_lm_objs(idx, soft_obj.name)
            
        else:
            # LANDMARKS NORMALES: hard + soft + línea
            # HARD = rojo
            hard_obj = create_landmark_sphere(
                f"{nice_name}_HARD",
                p_hard,
                1.0,
                (1, 0, 0, 1),
                cols["hard"].name
            )
            hard_obj.hide_select = False
            hard_obj.lock_location = (False, False, False)
            lm.hard_obj_name = hard_obj.name

            # SOFT = azul
            soft_obj = create_landmark_sphere(
                f"{nice_name}_SOFT",
                p_soft,
                1.0,
                (0, 0.6, 1, 1),
                cols["soft"].name
            )
            soft_obj.hide_select = False
            soft_obj.lock_location = (False, False, False)
            _remove_soft_surface_constraint(soft_obj)
            lm.soft_obj_name = soft_obj.name

            # LINEA
            line_obj = create_line_object(
                name=f"{nice_name}_LINE",
                p0=p_hard,
                p1=p_soft,
                hard_obj=hard_obj,
                soft_obj=soft_obj,
            )
            lm.line_obj_name = line_obj.name
            line_obj["tesis_lm_index"] = int(idx)
            line_obj["tesis_lm_role"] = "LINE"
            _register_lm_objs(idx, soft_obj.name, hard_obj.name)

        perf["create objects"] = time.perf_counter() - objects_started
        perf["TOTAL"] = time.perf_counter() - landmark_started
        _print_perf_block("LANDMARK", perf)
        # marcar que hay cambios sin guardar
        state.dirty = True
        _update_lm_counts(state)

        # Avanzar solo durante el flujo secuencial. Una corrección individual
        # debe quedarse en el landmark seleccionado.
        if self.sequential:
            next_idx = _next_sequential_idx(state.active_landmark)
            if next_idx is not None:
                state.active_landmark = next_idx

        return True


class TESIS_OT_RefreshBVH(Operator):
    bl_idname = "tesis.refresh_bvh"
    bl_label = "Refrescar BVH"
    bl_description = "Actualiza la caché BVH de las mallas (necesario después de ediciones de geometría)"

    def execute(self, context):
        state = context.scene.tesis_state
        skull, soft = obtener_mallas_paciente()
        if not skull or not soft:
            self.report({'WARNING'}, "Importa mallas primero.")
            return {'CANCELLED'}

        invalidate_bvh_cache()
        ensure_bvh_cache(context, skull, soft)

        self.report({'INFO'}, "BVH actualizado.")
        return {'FINISHED'}


def _draw_quick_tools(layout, context):
    """Barra compacta para las acciones frecuentes durante la medición."""
    state = context.scene.tesis_state
    row = layout.row(align=True)
    row.scale_y = 0.90
    row.operator("tesis.toggle_hard_vis", text="H", icon='BONE_DATA')
    row.operator("tesis.toggle_soft_vis", text="S", icon='MESH_UVSPHERE')
    row.operator("tesis.toggle_lines_vis", text="L", icon='OUTLINER_DATA_CURVE')
    row.operator("tesis.toggle_lm_names", text="Aa", icon='FONT_DATA')
    row.operator("tesis.refresh_bvh", text="", icon='FILE_REFRESH')

    trash = row.row(align=True)
    obj = getattr(context.view_layer.objects, "active", None)
    trash.enabled = bool(obj is not None and _resolve_lm_entry_for_object(state, obj) is not None)
    trash.operator("tesis.delete_selected", text="", icon='TRASH')


class TESIS_OT_SaveCaseMetadata(Operator):
    bl_idname = "tesis.save_case_metadata"
    bl_label = "Guardar cambios del caso al Excel"
    bl_description = "Guarda los campos editados del caso (edad, sexo, grupo, peso, observaciones) de vuelta al Excel"

    _error = None

    @staticmethod
    def _save_sync(excel_path, paciente, ojo, edad, sexo, grupo, peso, obs, has_neck, faltan_lm):
        """Escribe metadatos de forma síncrona en el hilo principal de Blender."""
        try:
            with open_workbook(excel_path, save=True) as wb:
                ws = wb[SHEET_CASES]
                headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
                col_pac = headers.get(COL_PACIENTE, 1)
                col_ojo = headers.get(COL_OJO, 2)
                cases_index = _build_patient_row_index(ws, col_pac=col_pac, col_ojo=col_ojo)
                r = cases_index.get((paciente, ojo))
                if r:
                    ws.cell(r, headers[COL_EDAD]).value = edad
                    ws.cell(r, headers[COL_SEXO]).value = sexo
                    ws.cell(r, headers[COL_GRUPO]).value = grupo
                    ws.cell(r, headers[COL_PESO]).value = peso
                    ws.cell(r, headers[COL_OBS]).value = obs
                    if COL_CUELLO in headers:
                        _ws_set_value(ws, r, headers[COL_CUELLO], 1 if has_neck else 0)
                    if COL_FALTAN in headers:
                        _ws_set_value(ws, r, headers[COL_FALTAN], 1 if faltan_lm else 0)
                else:
                    raise ValueError(f"No se encontró la fila operativa de {paciente} / {ojo} en Casos.")
            TESIS_OT_SaveCaseMetadata._error = None
        except Exception as e:
            log.exception("Error guardando metadatos paciente=%s ruta=%s", paciente, excel_path)
            TESIS_OT_SaveCaseMetadata._error = e

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases:
            self.report({'WARNING'}, "No hay caso activo.")
            return {'CANCELLED'}

        case = state.cases[state.active_case]
        excel_path = get_excel_path()
        if not os.path.exists(excel_path):
            self.report({'ERROR'}, f"Excel no encontrado: {excel_path}")
            return {'CANCELLED'}

        args = (excel_path, case.paciente, case.ojo, case.edad, case.sexo,
                case.grupo, case.peso, case.obs, case.has_neck, case.faltan_landmarks)

        TESIS_OT_SaveCaseMetadata._error = None
        try:
            context.window.cursor_set('WAIT')
        except Exception:
            pass
        self.report({'INFO'}, "Guardando metadatos…")
        try:
            TESIS_OT_SaveCaseMetadata._save_sync(*args)
        finally:
            try:
                context.window.cursor_set('DEFAULT')
            except Exception:
                pass

        if TESIS_OT_SaveCaseMetadata._error:
            self.report({'ERROR'}, f"Error al guardar: {TESIS_OT_SaveCaseMetadata._error}")
            return {'CANCELLED'}

        self.report({'INFO'}, "Cambios guardados al Excel.")
        try:
            for area in context.screen.areas:
                area.tag_redraw()
        except Exception:
            pass
        return {'FINISHED'}


class TESIS_OT_ResetLandmark(Operator):
    bl_idname = "tesis.reset_landmark"
    bl_label = "Eliminar landmark"
    bl_description = "Elimina esferas y línea del landmark, vuelve a pendiente"
    bl_options = {'REGISTER', 'UNDO'}

    index: IntProperty(default=-1) # type: ignore

    @classmethod
    def description(cls, context, properties):
        idx = properties.index
        if idx >= 0 and idx < len(LANDMARKS_ESPERADOS):
            lid, _, lbl = LANDMARKS_ESPERADOS[idx]
            return f"Eliminar {lid} (borrar esferas y línea): {lbl}"
        return "Eliminar landmark (borrar esferas y línea)"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}

        idx = self.index
        if idx < 0:
            idx = state.active_landmark
        if idx < 0 or idx >= len(state.landmarks):
            return {'CANCELLED'}

        state.active_landmark = idx
        if delete_landmark_complete(state, idx):
            lm_id = state.landmarks[idx].name if hasattr(state.landmarks[idx], 'name') else "?"
            self.report({'INFO'}, f"Landmark {idx+1}.{lm_id} reseteado a pendiente.")
            return {'FINISHED'}
        return {'CANCELLED'}


class TESIS_OT_SetLandmarkStatus(Operator):
    bl_idname = "tesis.set_landmark_status"
    bl_label = "No disponible"
    bl_description = "Cambia el landmark facial activo entre pendiente y no disponible"
    bl_options = {'REGISTER', 'UNDO'}

    status: StringProperty(default="MISSING") # type: ignore
    index: IntProperty(default=-1) # type: ignore  # -1 = usar active_landmark

    @classmethod
    def description(cls, context, properties):
        state = getattr(getattr(context, "scene", None), "tesis_state", None)
        idx = properties.index
        if idx < 0 and state is not None:
            idx = int(getattr(state, "active_landmark", -1))
        if 0 <= idx < len(LANDMARKS_ESPERADOS):
            lid, _, lbl = LANDMARKS_ESPERADOS[idx]
            return f"Cambiar disponibilidad de {idx + 1}.{lid}: {lbl}"
        return "Cambiar disponibilidad del landmark activo"

    def _resolve_index(self, state):
        idx = self.index if self.index >= 0 else state.active_landmark
        if idx < 0 or idx >= len(state.landmarks):
            return -1
        return idx

    def invoke(self, context, event):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}

        idx = self._resolve_index(state)
        if idx < 0:
            return {'CANCELLED'}

        # Este control corresponde solo a los 26 landmarks faciales.
        if idx not in _OTHER_INDICES:
            self.report({'WARNING'}, "El cuello se gestiona desde su sección propia.")
            return {'CANCELLED'}

        lm = state.landmarks[idx]
        if self.status == "MISSING" and lm.status == "SET":
            return context.window_manager.invoke_confirm(
                self,
                event,
                message=(
                    f"{_ui_landmark_label(idx)} ya tiene una medición. "
                    "Marcarlo como no disponible eliminará HARD, SOFT y la línea "
                    "del caso actual. ¿Continuar?"
                ),
            )
        return self.execute(context)

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}

        idx = self._resolve_index(state)
        if idx < 0 or idx not in _OTHER_INDICES:
            return {'CANCELLED'}

        state.active_landmark = idx
        lm = state.landmarks[idx]

        # Si ya era MISSING, el mismo botón lo reactiva como pendiente.
        if self.status == "MISSING" and lm.status == "MISSING":
            lm.status = "OMIT"
            lm.hx = lm.hy = lm.hz = 0.0
            lm.x = lm.y = lm.z = 0.0
            lm.thickness = 0.0
            lm.hard_obj_name = ""
            lm.soft_obj_name = ""
            lm.line_obj_name = ""
            state.dirty = True
            _update_lm_counts(state)
            if _marking_modal_active:
                _update_marking_header_from_state(state)
            self.report({'INFO'}, f"{_ui_landmark_label(idx)} volvió a pendiente.")
            return {'FINISHED'}

        # SET -> MISSING: borrar primero la medición visual/lógica completa.
        if lm.status == "SET":
            delete_landmark_complete(state, idx)
            lm = state.landmarks[idx]

        # OMIT -> MISSING (o SET ya limpiado -> MISSING).
        clear_landmark_objects(idx)
        _unregister_lm_objs(idx)
        lm.status = self.status
        lm.hx = lm.hy = lm.hz = 0.0
        lm.x = lm.y = lm.z = 0.0
        lm.thickness = 0.0
        lm.hard_obj_name = ""
        lm.soft_obj_name = ""
        lm.line_obj_name = ""
        state.dirty = True
        _update_lm_counts(state)

        # Si el flujo secuencial está activo, avanzar igual que al medir.
        if self.status == "MISSING" and _marking_modal_active:
            next_idx = _next_sequential_idx(idx)
            if next_idx is not None:
                state.active_landmark = next_idx
            _update_marking_header_from_state(state)

        self.report({'INFO'}, f"{_ui_landmark_label(idx)} marcado como no disponible.")
        return {'FINISHED'}


class TESIS_OT_MarkLandmark(Operator):
    bl_idname = "tesis.mark_landmark"
    bl_label = "Marcar landmark"
    bl_description = "Click para marcar este landmark en la malla"

    lm_id: StringProperty(default="") # type: ignore

    @classmethod
    def description(cls, context, properties):
        """Tooltip dinámico: muestra la descripción anatómica del landmark."""
        target = normalize_landmark_name(properties.lm_id) if properties.lm_id else ""
        idx = _LM_NAME_TO_IDX.get(target)
        if idx is not None and idx < len(LANDMARKS_ESPERADOS):
            _lid, _sid, lbl = LANDMARKS_ESPERADOS[idx]
            return f"{_lid}: {lbl}"
        return "Marcar este landmark"

    def execute(self, context):
        state = context.scene.tesis_state

        if not state.landmarks:
            self.report({'ERROR'}, "No hay landmarks cargados.")
            return {'CANCELLED'}

        target = normalize_landmark_name(self.lm_id)
        if not target:
            self.report({'ERROR'}, "lm_id vacío.")
            return {'CANCELLED'}

        idx = None
        for i, lm in enumerate(state.landmarks):
            if normalize_landmark_name(getattr(lm, "name", "")) == target:
                idx = i
                break

        if idx is None:
            self.report({'ERROR'}, f"No encontré landmark: {self.lm_id}")
            return {'CANCELLED'}

        state.active_landmark = idx
        # Validar el nuevo click antes de reemplazar una medición existente.
        skull, soft = obtener_mallas_paciente(quiet=True)
        if skull is None or soft is None:
            self.report({'WARNING'}, "Importa las mallas del caso activo antes de marcar landmarks.")
            return {'CANCELLED'}

        # Si ya estamos en VIEW_3D, usar contexto actual.
        if getattr(context, "area", None) and context.area.type == 'VIEW_3D':
            try:
                return bpy.ops.tesis.add_landmark('INVOKE_DEFAULT', sequential=True, start_from_active=True)
            except RuntimeError:
                self.report({'ERROR'}, "No se pudo iniciar el marcado en esta vista.")
                return {'CANCELLED'}

        # Si el botón se ejecuta desde otro editor, buscar una VIEW_3D y hacer override.
        screen = getattr(context, "screen", None)
        if screen is None:
            self.report({'ERROR'}, "No hay una vista 3D activa para iniciar marcado.")
            return {'CANCELLED'}

        for area in screen.areas:
            if area.type != 'VIEW_3D':
                continue
            region = next((r for r in area.regions if r.type == 'WINDOW'), None)
            if region is None:
                continue
            space = area.spaces.active
            try:
                with context.temp_override(area=area, region=region, space_data=space):
                    return bpy.ops.tesis.add_landmark('INVOKE_DEFAULT', sequential=True, start_from_active=True)
            except RuntimeError:
                continue

        self.report({'ERROR'}, "No encontré una VIEW_3D válida para iniciar marcado.")
        return {'CANCELLED'}


def _sync_landmark_data_from_objects(state, idx, mark_dirty=True):
    """Lee posiciones WORLD y actualiza un único LandmarkItem sin proyectarlo."""
    if state is None or idx < 0 or idx >= len(state.landmarks):
        return False
    lm = state.landmarks[idx]
    if lm.status != "SET":
        return False
    soft_obj = bpy.data.objects.get(getattr(lm, "soft_obj_name", ""))
    if soft_obj is None:
        return False
    p_soft = soft_obj.matrix_world.translation.copy()
    lm.x, lm.y, lm.z = p_soft.x, p_soft.y, p_soft.z
    if normalize_landmark_name(lm.name) in _NECK_IDS:
        lm.thickness = 0.0
    else:
        hard_obj = bpy.data.objects.get(getattr(lm, "hard_obj_name", ""))
        if hard_obj is None:
            return False
        p_hard = hard_obj.matrix_world.translation.copy()
        lm.hx, lm.hy, lm.hz = p_hard.x, p_hard.y, p_hard.z
        lm.thickness = (p_soft - p_hard).length
        line_obj = bpy.data.objects.get(getattr(lm, "line_obj_name", ""))
        if line_obj is not None:
            update_line_endpoints(line_obj, p_hard, p_soft)
    if mark_dirty:
        state.dirty = True
    return True


def _recalculate_single_landmark(context, lm, skull_soft=None, depsgraph=None):
    """Compatibilidad: sincroniza datos actuales sin cambiar posiciones."""
    state = context.scene.tesis_state
    idx = next((i for i, item in enumerate(state.landmarks) if item.name == lm.name), -1)
    return _sync_landmark_data_from_objects(state, idx)


class TESIS_OT_RecalculateActiveSOFT(bpy.types.Operator):
    bl_idname = "tesis.recalculate_active_soft"
    bl_label = "Recalcular activo"
    bl_description = "Sincroniza el landmark activo y actualiza su grosor"
    bl_options = {'REGISTER', 'UNDO'}

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            self.report({'WARNING'}, "No hay landmarks cargados.")
            return {'CANCELLED'}

        idx = max(0, min(state.active_landmark, len(state.landmarks) - 1))
        lm = state.landmarks[idx]
        if lm.status != "SET":
            self.report({'WARNING'}, "El landmark activo no está marcado.")
            return {'CANCELLED'}

        skull_soft = obtener_mallas_paciente()
        if skull_soft[0] is None and skull_soft[1] is None:
            self.report({'WARNING'}, "No se encontraron mallas del caso. Importa primero.")
            return {'CANCELLED'}

        ok = _recalculate_single_landmark(context, lm, skull_soft=skull_soft, depsgraph=context.evaluated_depsgraph_get())
        if not ok:
            self.report({'ERROR'}, f"No se pudo recalcular {lm.name}.")
            return {'CANCELLED'}

        state.dirty = True
        for area in context.screen.areas:
            area.tag_redraw()
        self.report({'INFO'}, f"{lm.name}: {lm.thickness:.2f} mm")
        return {'FINISHED'}


class TESIS_OT_RecalculateAllSOFT(bpy.types.Operator):
    bl_idname = "tesis.recalculate_all_soft"
    bl_label = "Recalcular TODOS los SOFT"
    bl_description = "Sincroniza los landmarks existentes y recalcula sus grosores"
    bl_options = {'REGISTER', 'UNDO'}

    def execute(self, context):
        state = context.scene.tesis_state

        if not state.landmarks:
            self.report({'WARNING'}, "No hay landmarks cargados.")
            return {'CANCELLED'}

        # Obtener mallas y depsgraph UNA vez para todo el lote
        skull_soft = obtener_mallas_paciente()
        if skull_soft[0] is None and skull_soft[1] is None:
            self.report({'WARNING'}, "No se encontraron mallas del caso. Importa primero.")
            return {'CANCELLED'}
        depsgraph = context.evaluated_depsgraph_get()

        updated = 0
        errors = 0
        for lm in state.landmarks:
            if lm.status != "SET":
                continue
            ok = _recalculate_single_landmark(context, lm, skull_soft=skull_soft, depsgraph=depsgraph)
            if ok:
                updated += 1
            else:
                errors += 1

        state.dirty = True
        for area in context.screen.areas:
            area.tag_redraw()
        msg = f"Recalculados: {updated} landmarks"
        if errors:
            msg += f" ({errors} errores)"
        self.report({'INFO'}, msg)
        return {'FINISHED'}


# =============================================================================
# MEDICIÓN DE CUELLO (CCU)
# =============================================================================
# El estudio contempla una medición secundaria del cuello con la CCU, usando
# landmarks específicos del cuello (anterior y posterior). El proceso genera un
# contorno alrededor de la zona cervical a partir de una proyección radial sobre
# la piel externa.
#
# Este cálculo se mantiene separado del grosor facial y se documenta como una
# variable complementaria de la anatomía regional del cuello.
# =============================================================================
class TESIS_OT_MarkNeckPoint(Operator):
    """Marca un único punto cervical de manera independiente.

    neck_ant y neck_post se editan por separado. Cada invocación espera UN clic
    sobre SOFT y termina inmediatamente. ESC/RMB cancela únicamente ese punto y
    restaura su estado anterior.
    """
    bl_idname = "tesis.mark_neck_point"
    bl_label = "Marcar punto cervical"
    bl_description = "Marca un único punto cervical sobre la superficie blanda. ESC cancela solo este punto."
    bl_options = {'REGISTER'}

    target: StringProperty(default="neck_ant") # type: ignore

    def invoke(self, context, event):
        state = context.scene.tesis_state

        if not state.cases:
            self.report({'WARNING'}, "No hay caso activo.")
            return {'CANCELLED'}

        case = state.cases[state.active_case]
        if not getattr(case, "has_neck", False):
            self.report({'WARNING'}, "Este caso no tiene cuello habilitado.")
            return {'CANCELLED'}

        target = normalize_landmark_name(self.target)
        if target not in _NECK_IDS:
            self.report({'ERROR'}, f"Punto cervical inválido: {self.target}")
            return {'CANCELLED'}

        lm = find_landmark_by_name(state.landmarks, self.target)
        idx = _LM_NAME_TO_IDX.get(_normalize_lm(self.target))
        if lm is None or idx is None:
            self.report({'ERROR'}, f"No existe el landmark {self.target}.")
            return {'CANCELLED'}

        skull, soft = obtener_mallas_paciente(quiet=True)
        if skull is None or soft is None:
            self.report({'WARNING'}, "Importa las mallas del caso antes de marcar el cuello.")
            return {'CANCELLED'}

        ensure_bvh_cache(context, skull, soft)

        # Guardar SOLO el punto que se está editando para poder restaurarlo con ESC.
        original_obj = bpy.data.objects.get(getattr(lm, "soft_obj_name", ""))
        visual = None
        if original_obj is not None:
            visual = {
                "location": original_obj.matrix_world.translation.copy(),
                "scale": original_obj.scale.copy(),
                "hide_viewport": bool(original_obj.hide_viewport),
                "hide_render": bool(original_obj.hide_render),
                "hide_select": bool(original_obj.hide_select),
            }

        self._original = {
            "status": lm.status,
            "x": float(lm.x), "y": float(lm.y), "z": float(lm.z),
            "hx": float(lm.hx), "hy": float(lm.hy), "hz": float(lm.hz),
            "thickness": float(lm.thickness),
            "soft_obj_state": visual,
            "dirty": bool(state.dirty),
        }
        self._idx = int(idx)

        context.window_manager.modal_handler_add(self)

        readable = "anterior" if target == "NECK_ANT" else "posterior"
        if context.area:
            context.area.header_text_set(
                f"CUELLO · marque el punto {readable} · ESC cancela"
            )
        self.report({'INFO'}, f"Marque el punto cervical {readable}.")
        return {'RUNNING_MODAL'}

    def _restore_original(self, context):
        """Restaura exclusivamente el punto cervical editado."""
        if not hasattr(self, "_original") or not hasattr(self, "_idx"):
            return

        state = context.scene.tesis_state
        idx = int(self._idx)
        if idx < 0 or idx >= len(state.landmarks):
            return

        lm = state.landmarks[idx]
        orig = self._original

        clear_landmark_objects(idx)
        _unregister_lm_objs(idx)

        lm.status = orig.get("status", "OMIT")
        lm.x = orig.get("x", 0.0)
        lm.y = orig.get("y", 0.0)
        lm.z = orig.get("z", 0.0)
        lm.hx = orig.get("hx", 0.0)
        lm.hy = orig.get("hy", 0.0)
        lm.hz = orig.get("hz", 0.0)
        lm.thickness = orig.get("thickness", 0.0)
        lm.soft_obj_name = ""
        lm.hard_obj_name = ""
        lm.line_obj_name = ""

        if lm.status == "SET":
            visual = orig.get("soft_obj_state")
            location = (
                visual["location"]
                if visual is not None
                else Vector((lm.x, lm.y, lm.z))
            )
            radius = float(visual["scale"].x) if visual is not None else 1.5
            cols = ensure_tesis_collections()
            restored = create_landmark_sphere(
                f"{idx:02d}_{lm.name}_SOFT",
                location,
                radius,
                (0.15, 0.85, 0.30, 1.0),
                cols["soft"].name,
            )
            if visual is not None:
                restored.hide_viewport = visual["hide_viewport"]
                restored.hide_render = visual["hide_render"]
                restored.hide_select = visual["hide_select"]
            else:
                restored.hide_select = False
            restored.lock_location = (False, False, False)
            lm.soft_obj_name = restored.name
            _register_lm_objs(idx, restored.name)

        state.dirty = bool(orig.get("dirty", False))
        _update_lm_counts(state)
        _tag_tesis_ui_redraw(context)

    def _raycast_soft(self, origin, direction):
        """Marca cuello directamente sobre la malla SOFT completa visible.

        El filtro de componente principal se usa para reducir placas en el
        automatismo facial, pero NO participa en neck_ant / neck_post ni CCU.
        """
        full_bvh = _BVH_CACHE.get("soft")
        if full_bvh is None:
            return None
        return full_bvh.ray_cast(origin, direction)[0]

    def modal(self, context, event):
        if event.type in {'ESC', 'RIGHTMOUSE'} and event.value == 'PRESS':
            self._restore_original(context)
            if context.area:
                context.area.header_text_set(None)
            self.report({'INFO'}, "Marcado cervical cancelado.")
            return {'CANCELLED'}

        if event.type == 'LEFTMOUSE' and event.value in {'PRESS', 'CLICK'}:
            if context.area is None or context.area.type != 'VIEW_3D':
                return {'PASS_THROUGH'}

            region = next((r for r in context.area.regions if r.type == 'WINDOW'), None)
            rv3d = context.space_data.region_3d if context.space_data else None
            if region is None or rv3d is None:
                return {'PASS_THROUGH'}

            coord_x = event.mouse_x - region.x
            coord_y = event.mouse_y - region.y
            if not (0 <= coord_x < region.width and 0 <= coord_y < region.height):
                return {'RUNNING_MODAL'}

            coord = (coord_x, coord_y)
            origin = view3d_utils.region_2d_to_origin_3d(region, rv3d, coord)
            direction = view3d_utils.region_2d_to_vector_3d(
                region, rv3d, coord
            ).normalized()

            point = self._raycast_soft(origin, direction)
            if point is None:
                self.report({'WARNING'}, "No intersectó la superficie blanda visible.")
                return {'RUNNING_MODAL'}

            state = context.scene.tesis_state
            idx = int(self._idx)
            lm = state.landmarks[idx]

            lm.status = "SET"
            lm.x, lm.y, lm.z = point.x, point.y, point.z
            lm.hx = lm.hy = lm.hz = 0.0
            lm.thickness = 0.0

            clear_landmark_objects(idx)
            _unregister_lm_objs(idx)

            cols = ensure_tesis_collections()
            sphere = create_landmark_sphere(
                f"{idx:02d}_{lm.name}_SOFT",
                point,
                1.5,
                (0.15, 0.85, 0.30, 1.0),
                cols["soft"].name,
            )
            sphere.hide_select = False
            sphere.lock_location = (False, False, False)
            lm.soft_obj_name = sphere.name
            lm.hard_obj_name = ""
            lm.line_obj_name = ""
            _register_lm_objs(idx, sphere.name)

            # Al cambiar UNA referencia cervical, el anillo visual previo deja
            # de representar esos puntos. Se elimina solo el anillo; CCU_mm
            # histórica permanece intacta.
            case = state.cases[state.active_case]
            _, soft_obj = obtener_mallas_paciente(quiet=True)
            _remove_current_ccu_ring(case, soft_obj)

            state.dirty = True
            _update_lm_counts(state)

            if context.area:
                context.area.header_text_set(None)
            _tag_tesis_ui_redraw(context)

            readable = "anterior" if normalize_landmark_name(lm.name) == "NECK_ANT" else "posterior"
            if float(getattr(case, "ccu_mm", 0.0)) > 0:
                self.report(
                    {'WARNING'},
                    f"Punto {readable} actualizado. La CCU guardada no cambió; regenere el anillo cuando ambos puntos estén correctos."
                )
            else:
                self.report({'INFO'}, f"Punto cervical {readable} marcado.")
            return {'FINISHED'}

        if event.type in {
            'MIDDLEMOUSE', 'WHEELUPMOUSE', 'WHEELDOWNMOUSE',
            'TRACKPADPAN', 'TRACKPADZOOM'
        }:
            return {'PASS_THROUGH'}

        return {'RUNNING_MODAL'}

    def cancel(self, context):
        self._restore_original(context)
        if context.area:
            context.area.header_text_set(None)

class TESIS_OT_CreateCCURing(bpy.types.Operator):
    bl_idname = "tesis.create_ccu_ring"
    bl_label = "Crear / regenerar anillo CCU"
    bl_description = "Genera el contorno cervical sobre la malla SOFT completa con el método radial anclado"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases:
            self.report({'WARNING'}, "No hay caso activo.")
            return {'CANCELLED'}

        case = state.cases[state.active_case]
        if not getattr(case, "has_neck", False):
            self.report({'WARNING'}, "Este caso no tiene cuello.")
            return {'CANCELLED'}

        skull, soft_obj = obtener_mallas_paciente(quiet=True)
        if skull is None or soft_obj is None:
            self.report({'ERROR'}, "No hay mallas importadas para el caso.")
            return {'CANCELLED'}

        # Normalmente será CACHE HIT. No se construye una BVH específica para CCU.
        ensure_bvh_cache(context, skull, soft_obj)
        ring = build_ccu_ring_geometry(
            context, case, state, soft_obj, reconstruction=False
        )
        if ring is None:
            self.report({'WARNING'}, "No fue posible construir el anillo CCU sobre la superficie disponible.")
            return {'CANCELLED'}

        estimated = int(ring.get("ccu_interpolated_points", 0))
        if estimated:
            self.report(
                {'WARNING'},
                f"Anillo creado: {estimated}/256 sectores fueron estimados entre superficies válidas."
            )
        else:
            self.report({'INFO'}, "Anillo CCU creado sobre la superficie cutánea.")
        return {'FINISHED'}



class TESIS_OT_ConfirmCCU(bpy.types.Operator):
    bl_idname = "tesis.confirm_ccu"
    bl_label = "Confirmar CCU"
    bl_description = "Registra como CCU la longitud del anillo visual actualmente revisado"

    def invoke(self, context, event):
        state = context.scene.tesis_state
        if state.cases:
            case = state.cases[state.active_case]
            _, soft_obj = obtener_mallas_paciente(quiet=True)
            ring = _get_current_ccu_ring(case, soft_obj)
            new_value = _compute_ring_length_mm(ring)
            if float(getattr(case, "ccu_mm", 0.0)) > 0 and new_value > 0:
                return context.window_manager.invoke_confirm(
                    self,
                    event,
                    message=(
                        f"Ya existe una CCU guardada de {case.ccu_mm:.1f} mm. "
                        f"¿Reemplazarla por {new_value:.1f} mm?"
                    ),
                )
        return self.execute(context)

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases:
            return {'CANCELLED'}

        case = state.cases[state.active_case]
        _, soft_obj = obtener_mallas_paciente(quiet=True)
        if soft_obj is None:
            self.report({'ERROR'}, "No se encontró malla SOFT.")
            return {'CANCELLED'}

        ring = _get_current_ccu_ring(case, soft_obj)
        if ring is None:
            self.report({'ERROR'}, "No existe anillo CCU. Créelo o regenérelo primero.")
            return {'CANCELLED'}

        interpolated_points = int(ring.get("ccu_interpolated_points", 0))
        total_points = max(1, len(getattr(ring.data, "vertices", [])))
        interpolated_ratio = interpolated_points / total_points
        if interpolated_ratio > 0.25:
            self.report(
                {'ERROR'},
                f"CCU no confirmada: {interpolated_points}/{total_points} puntos fueron estimados."
            )
            return {'CANCELLED'}
        if interpolated_points > 0:
            self.report(
                {'WARNING'},
                f"CCU con {interpolated_points}/{total_points} puntos interpolados; revisar visualmente."
            )

        total_length = _compute_ring_length_mm(ring)
        num_edges = len(getattr(ring.data, "edges", []))
        if total_length <= 0:
            self.report({'ERROR'}, "No fue posible calcular la longitud del anillo CCU.")
            return {'CANCELLED'}
        if num_edges < 200:
            self.report(
                {'WARNING'},
                f"Anillo CCU tiene solo {num_edges} segmentos (esperado: 256). Calidad reducida."
            )

        # Solo esta acción reemplaza el valor CCU del caso.
        case.ccu_mm = float(total_length)
        if total_length < 200 or total_length > 500:
            self.report(
                {'WARNING'},
                f"CCU = {total_length:.1f} mm fuera del rango operativo 200–500 mm; revisar geometría."
            )

        case.estado_nutricional = estado_nutricional(case.sexo, case.edad, case.ccu_mm)
        ring["ccu_saved_length_mm"] = float(case.ccu_mm)
        ring["ccu_reconstructed"] = False
        ring.display_type = 'WIRE'
        ring.hide_select = True
        state.dirty = True

        self.report(
            {'INFO'},
            f"CCU = {total_length:.2f} mm | Estado: {case.estado_nutricional}"
        )
        return {'FINISHED'}



def _sync_scene_landmarks_before_save(context, state):
    """Toma el snapshot actual sin reproyectar landmarks manuales."""
    return _sync_landmark_state_from_evaluated(
        context.scene, context.evaluated_depsgraph_get(), mark_dirty=False
    )


_SAVE_IN_PROGRESS = False

# =============================================================================
# GUARDADO DEL CASO
# =============================================================================
# El guardado es la etapa final del flujo de medición. Es aquí donde se escribe
# el estado del caso activo en el Excel y donde se actualizan los resúmenes del
# proyecto.
#
# El programa permite guardar avances intermedios y, cuando se cumplen las
# condiciones del protocolo, también marca el caso como procesado. Esto separa
# el progreso temporal de la finalización del caso.
# =============================================================================
class TESIS_OT_SaveCase(Operator):
    bl_idname = "tesis.save_case"
    bl_label = "Guardar en Excel (Landmarks)"
    bl_description = "Guarda todos los landmarks del caso actual al Excel (hoja Landmarks y hoja Resumen)"

    advance_after_save: BoolProperty(default=False) # type: ignore
    finalize_case: BoolProperty(default=False) # type: ignore

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self._error = None
        self._timer = None
        self._snap_case = None
        self._snap_lm = None
        self._snap_unit = 1.0
        self._snap_mark = False
        self._snap_summary = None
        self._snap_excel_path = ""
        self._previous_processed = False
        self._saved_case_index = -1
        self._saved_case_key = ""
        self._next_import_pending = False
        self._error_mode = False

    def _snapshot(self, context):
        """Captura un snapshot estable antes de escribir el Excel."""
        state = context.scene.tesis_state
        case = state.cases[state.active_case]
        self._error_mode = (_safe_str(getattr(state, "measurement_mode", "NORMAL")).upper() == "ERROR")
        if (not self._error_mode
                and _safe_str(getattr(state, "session_id", _PRIMARY_MEASUREMENT_SESSION)).upper() != _PRIMARY_MEASUREMENT_SESSION):
            raise ValueError(
                f"Modo NORMAL solo puede guardar la medición principal como {_PRIMARY_MEASUREMENT_SESSION}. "
                "Para S2 u otra sesión usa modo ERROR; para re-revisar el principal, vuelve a S1."
            )
        _sync_scene_landmarks_before_save(context, state)
        lm_data = _serialize_landmarks(state.landmarks)
        case_data = _serialize_case(case, state)
        if self._error_mode and not bool(getattr(state, "error_include_neck", True)):
            case_data["ccu_mm"] = 0.0
            case_data["estado_nutricional"] = "No aplicable"
            for idx in _NECK_INDICES:
                if idx < len(lm_data):
                    lm_data[idx]["status"] = "OMIT"
                    lm_data[idx]["x"] = lm_data[idx]["y"] = lm_data[idx]["z"] = None
                    lm_data[idx]["hx"] = lm_data[idx]["hy"] = lm_data[idx]["hz"] = None
                    lm_data[idx]["thickness"] = None
        unit_scale = bpy.context.scene.unit_settings.scale_length
        self._snap_excel_path = get_excel_path()
        self._previous_processed = bool(case.procesado)
        self._saved_case_index = state.active_case
        self._saved_case_key = _current_case_key(case)

        neck_states = [
            lm_data[idx].get("status", "OMIT") if idx < len(lm_data) else "OMIT"
            for idx in _NECK_INDICES
        ]
        neck_complete = len(neck_states) == 2 and all(status == "SET" for status in neck_states)
        neck_empty = not any(status == "SET" for status in neck_states)
        if not case_data.get("has_neck"):
            case_data["has_neck"] = False
            case_data["ccu_mm"] = 0.0
            case_data["estado_nutricional"] = "No aplicable"
            for idx in _NECK_INDICES:
                if idx < len(lm_data):
                    lm_data[idx]["status"] = "OMIT"
                    lm_data[idx]["x"] = 0.0
                    lm_data[idx]["y"] = 0.0
                    lm_data[idx]["z"] = 0.0
                    lm_data[idx]["hx"] = 0.0
                    lm_data[idx]["hy"] = 0.0
                    lm_data[idx]["hz"] = 0.0
                    lm_data[idx]["thickness"] = 0.0
        case_data["faltan_landmarks"] = any(
            idx < len(lm_data) and lm_data[idx].get("status") == "MISSING"
            for idx in _OTHER_INDICES
        )

        facial_complete = all(
            idx < len(lm_data) and lm_data[idx].get("status") in {"SET", "MISSING"}
            for idx in _OTHER_INDICES
        )
        if self.finalize_case:
            blockers = _case_finalize_blockers(state, case)
            if blockers:
                raise ValueError("No se puede finalizar: " + " ".join(blockers))

        th_vals = [d["thickness"] for d in lm_data
                   if d["status"] == "SET" and d["thickness"] and d["thickness"] > 0]
        summary = {
            "set": sum(1 for idx in _OTHER_INDICES if idx < len(lm_data) and lm_data[idx]["status"] == "SET"),
            "omit": sum(1 for idx in _OTHER_INDICES if idx < len(lm_data) and lm_data[idx]["status"] == "MISSING"),
            "pend": sum(1 for idx in _OTHER_INDICES if idx < len(lm_data) and lm_data[idx]["status"] == "OMIT"),
            **_thickness_summary_from_values(th_vals),
        }

        self._snap_case = case_data
        self._snap_lm = lm_data
        self._snap_unit = unit_scale
        self._snap_mark = bool(
            (not self._error_mode)
            and self.finalize_case and facial_complete
            and (not case_data.get("has_neck") or neck_complete)
        )
        self._snap_summary = summary

    def _save_excel_sync(self):
        """Escribe el snapshot de forma síncrona en el hilo principal de Blender."""
        try:
            if self._error_mode:
                save_error_measurement_to_excel(
                    self._snap_excel_path,
                    self._snap_case,
                    self._snap_lm,
                )
            else:
                save_landmarks_to_excel(
                    self._snap_excel_path,
                    self._snap_case,
                    self._snap_lm,
                    self._snap_unit,
                    mark_processed=self._snap_mark
                )
            self._error = None
        except Exception as e:
            log.exception(
                "Error guardando paciente=%s fase=save ruta=%s",
                (self._snap_case or {}).get("paciente", "?"), self._snap_excel_path,
            )
            self._error = e

    def execute(self, context):
        global _SAVE_IN_PROGRESS
        state = context.scene.tesis_state
        if not state.cases:
            return {'CANCELLED'}
        if state.cases[state.active_case].coordinate_review_required:
            self.report({'WARNING'}, "Este caso está en modo revisión de coordenadas. No se guardará para evitar sobrescribir datos históricos.")
            return {'CANCELLED'}
        if not state.cases[state.active_case].coordinate_frame_valid:
            self.report({'ERROR'}, "Coordenadas incompatibles: no se modificó el Excel.")
            return {'CANCELLED'}
        if _SAVE_IN_PROGRESS:
            self.report({'WARNING'}, "Ya hay un guardado en curso.")
            return {'CANCELLED'}

        # Capturar snapshot en el hilo principal
        try:
            self._snapshot(context)
        except ValueError as exc:
            self.report({'WARNING'}, str(exc))
            return {'CANCELLED'}

        # Escritura segura: el Excel termina de guardarse antes de que Blender
        # cambie de caso o vuelva a manipular la escena. No se crean threads.
        self._error = None
        _SAVE_IN_PROGRESS = True
        try:
            context.window.cursor_set('WAIT')
        except Exception:
            pass
        self.report({'INFO'}, "Guardando landmarks…")
        self._save_excel_sync()

        # Se conserva un único ciclo modal para aplicar el resultado y, en
        # Guardar→Siguiente, importar el próximo caso en un ciclo posterior.
        # Este timer corre en el hilo principal y no hace I/O concurrente.
        self._timer = context.window_manager.event_timer_add(0.05, window=context.window)
        context.window_manager.modal_handler_add(self)
        return {'RUNNING_MODAL'}

    def modal(self, context, event):
        global _SAVE_IN_PROGRESS
        if event.type != 'TIMER':
            return {'PASS_THROUGH'}
        if self._next_import_pending:
            self._next_import_pending = False
            context.window_manager.event_timer_remove(self._timer)
            self._timer = None
            state = context.scene.tesis_state
            if state.cases and state.active_case < len(state.cases):
                next_case = state.cases[state.active_case]
                if next_case.ruta_hueso.strip() and next_case.ruta_blando.strip():
                    bpy.ops.tesis.import_meshes('EXEC_DEFAULT')
            for area in context.screen.areas:
                area.tag_redraw()
            return {'FINISHED'}
        try:
            context.window.cursor_set('DEFAULT')
        except Exception:
            pass

        if self._error:
            context.window_manager.event_timer_remove(self._timer)
            self._timer = None
            _SAVE_IN_PROGRESS = False
            self.report({'ERROR'}, str(self._error))
            return {'CANCELLED'}

        state = context.scene.tesis_state
        if self._saved_case_index < 0 or self._saved_case_index >= len(state.cases):
            context.window_manager.event_timer_remove(self._timer)
            self._timer = None
            _SAVE_IN_PROGRESS = False
            self.report({'ERROR'}, "El caso original ya no está disponible; no se aplicó el resumen.")
            return {'CANCELLED'}
        case = state.cases[self._saved_case_index]
        if _current_case_key(case) != self._saved_case_key:
            context.window_manager.event_timer_remove(self._timer)
            self._timer = None
            _SAVE_IN_PROGRESS = False
            self.report({'ERROR'}, "El caso activo cambió durante el guardado; no se aplicó el resumen.")
            return {'CANCELLED'}

        if self._error_mode:
            # Actualizar progreso en memoria sin tocar resúmenes ni datos canónicos.
            key = _error_progress_key_values(
                (self._snap_case or {}).get("case_id", case.paciente),
                (self._snap_case or {}).get("observer_id", "OBS1"),
                (self._snap_case or {}).get("session_id", "S1"),
                (self._snap_case or {}).get("replicate_id", 1),
                (self._snap_case or {}).get("error_cohort", ""),
            )
            entry = {"statuses": {}, "ccu_mm": _safe_float((self._snap_case or {}).get("ccu_mm"), 0.0)}
            for idx, lm in enumerate(self._snap_lm or []):
                if idx < len(LANDMARKS_ESPERADOS):
                    entry["statuses"][LANDMARKS_ESPERADOS[idx][0]] = _safe_str(lm.get("status", "OMIT")).upper()
            _ERROR_PROGRESS_CACHE[key] = entry
            state.dirty = False
            _SAVE_IN_PROGRESS = False

            if self.advance_after_save:
                if _advance_to_next_filtered_case(context, skip_processed=False):
                    self._next_import_pending = True
                    self.report({'INFO'}, "Medición de error guardada. Preparando el siguiente caso de la cohorte.")
                    return {'RUNNING_MODAL'}
                self.report({'INFO'}, "Medición de error guardada. No quedan casos pendientes en esta sesión.")

            context.window_manager.event_timer_remove(self._timer)
            self._timer = None
            self.report({'INFO'}, "Medición de error guardada en Mediciones_LM sin sobrescribir la medición principal.")
            for area in context.screen.areas:
                area.tag_redraw()
            return {'FINISHED'}

        s = self._snap_summary or {"set": 0, "omit": 0, "pend": 0, "avg": 0.0, "min": 0.0, "max": 0.0}
        case.summary_set = int(s["set"])
        case.summary_missing = int(s["omit"])
        case.summary_pend = int(s["pend"])
        case.summary_n_valid = int(s.get("n_valid", 0))
        case.summary_avg = float(s["avg"])
        case.summary_median = float(s.get("median", 0.0))
        case.summary_sd = float(s.get("sd", 0.0))
        case.summary_min = float(s["min"])
        case.summary_max = float(s["max"])
        case.summary_loaded = True
        snap_lm = self._snap_lm or []
        case.summary_neck_ant_status = str(
            snap_lm[_NECK_INDICES[0]].get("status", "OMIT")
            if len(snap_lm) > _NECK_INDICES[0] else "OMIT"
        ).upper()
        case.summary_neck_post_status = str(
            snap_lm[_NECK_INDICES[1]].get("status", "OMIT")
            if len(snap_lm) > _NECK_INDICES[1] else "OMIT"
        ).upper()
        snap_case = self._snap_case or {}
        global _CASE_PROPERTY_UPDATE_SUPPRESS
        _CASE_PROPERTY_UPDATE_SUPPRESS = True
        try:
            case.has_neck = bool(snap_case.get("has_neck", case.has_neck))
            case.ccu_mm = _safe_float(snap_case.get("ccu_mm", case.ccu_mm), case.ccu_mm)
            case.baseline_ccu_mm = case.ccu_mm
            case.baseline_estado_nutricional = _safe_str(snap_case.get("estado_nutricional", case.estado_nutricional))
            case.faltan_landmarks = bool(snap_case.get("faltan_landmarks", False))
            case.obs = _safe_str(snap_case.get("obs", case.obs))
        finally:
            _CASE_PROPERTY_UPDATE_SUPPRESS = False

        # Si el guardado confirmó que el caso no tiene cuello, reflejar EN MAIN THREAD
        # exactamente lo escrito por el snapshot y retirar solo los visuales cervicales.
        if not case.has_neck:
            _, soft_obj = obtener_mallas_paciente(quiet=True)
            _remove_current_ccu_ring(case=case, soft_obj=soft_obj)
            for idx in _NECK_INDICES:
                if idx < len(state.landmarks):
                    clear_landmark_objects(idx)
                    lm = state.landmarks[idx]
                    lm.status = "OMIT"
                    lm.hx = lm.hy = lm.hz = 0.0
                    lm.x = lm.y = lm.z = 0.0
                    lm.thickness = 0.0
                    lm.hard_obj_name = ""
                    lm.soft_obj_name = ""
                    lm.line_obj_name = ""
            case.ccu_mm = 0.0
            case.estado_nutricional = "No aplicable"
        if hasattr(case, "procesado"):
            case.procesado = True if self._snap_mark else self._previous_processed
            _update_case_counts(state)
        if hasattr(state, "dirty"):
            state.dirty = False

        _SAVE_IN_PROGRESS = False
        if self.advance_after_save:
            if _advance_to_next_filtered_case(context, skip_processed=True):
                self._next_import_pending = True
                self.report({'INFO'}, "Guardado confirmado. Preparando el siguiente caso.")
                return {'RUNNING_MODAL'}
            else:
                self.report({'INFO'}, "Guardado. No quedan casos que cumplan los filtros.")

        context.window_manager.event_timer_remove(self._timer)
        self._timer = None
        self.report({'INFO'}, "Guardado en Excel (hoja Coordenadas).")
        for area in context.screen.areas:
            area.tag_redraw()
        return {'FINISHED'}

    def invoke(self, context, event):
        """Validación pre-guardado."""
        state = context.scene.tesis_state
        if state.cases and state.cases[state.active_case].coordinate_review_required:
            self.report({'WARNING'}, "Este caso está en modo revisión de coordenadas. No se guardará para evitar sobrescribir datos históricos.")
            return {'CANCELLED'}
        if state.cases and not state.cases[state.active_case].coordinate_frame_valid:
            self.report({'ERROR'}, "Coordenadas incompatibles: no se modificó el Excel.")
            return {'CANCELLED'}
        
        if not state.landmarks:
            self.report({'WARNING'}, "No hay landmarks para guardar.")
            return {'CANCELLED'}
        
        # Usar contadores cacheados (ya actualizados por _update_lm_counts)
        set_count = state.lm_count_set
        pending_count = state.lm_count_pending
        
        if self.finalize_case:
            return self.execute(context)
        if set_count == 0:
            return context.window_manager.invoke_confirm(
                self, event, 
                message=f"No hay landmarks SET. \u00bfGuardar de todos modos?"
            )
        
        if pending_count > 0:
            return context.window_manager.invoke_confirm(
                self, event,
                message=f"Quedan {pending_count} landmarks PENDIENTES. \u00bfContinuar guardado?"
            )
        
        return self.execute(context)


class TESIS_OT_ClearCurrentCase(Operator):
    bl_idname = "tesis.clear_current_case"
    bl_label = "Borrar LM"
    bl_description = "Marca todos los landmarks del caso como pendientes. Requiere guardar para escribir el cambio en Excel."
    bl_options = {'REGISTER', 'UNDO'}

    def execute(self, context):
        state = context.scene.tesis_state
        removed = 0
        _OBJ_NAME_TO_LM.clear()

        # 1. Limpiar landmarks del caso actual sin eliminar las mallas importadas SKULL/SOFT.
        if state.landmarks:
            for i, lm in enumerate(state.landmarks):
                for attr in ("hard_obj_name", "soft_obj_name", "line_obj_name"):
                    obj_name = getattr(lm, attr, "")
                    if obj_name:
                        obj = bpy.data.objects.get(obj_name)
                        if obj:
                            _safe_remove_object(obj)
                            removed += 1
                clear_landmark_objects(i)
                lm.status = "OMIT"
                lm.thickness = 0.0
                lm.hx = lm.hy = lm.hz = 0.0
                lm.x = lm.y = lm.z = 0.0
                lm.hard_obj_name = ""
                lm.soft_obj_name = ""
                lm.line_obj_name = ""

        # 2. Limpiar únicamente el anillo CCU del caso activo.
        case = state.cases[state.active_case] if state.cases and state.active_case < len(state.cases) else None
        _, soft_obj = obtener_mallas_paciente(quiet=True)
        if _remove_current_ccu_ring(case=case, soft_obj=soft_obj):
            removed += 1

        for col_name in ("LM_HARD", "LM_SOFT", "LM_LINES", "TESIS_LANDMARKS", "TESIS_CCU", "TESIS_ROOT"):
            col = bpy.data.collections.get(col_name)
            if col and len(col.objects) == 0 and len(col.children) == 0:
                bpy.data.collections.remove(col)
        _cleanup_auxiliary_datablocks()

        if state.cases and state.active_case < len(state.cases):
            case = state.cases[state.active_case]
            for idx in _NECK_INDICES:
                if idx < len(state.landmarks):
                    state.landmarks[idx].status = "OMIT"
            case.ccu_mm = 0.0
            case.estado_nutricional = ""

        # SKULL/SOFT no cambian: conservar BVH evita reconstrucciones innecesarias.
        state.dirty = True
        state.active_landmark = 0
        _update_lm_counts(state)
        self.report({'INFO'}, f"LM borrados: {removed} objetos eliminados; las mallas SKULL/SOFT y sus BVH se conservan, guarda para escribir el cambio en Excel.")
        return {'FINISHED'}

    def invoke(self, context, event):
        return context.window_manager.invoke_confirm(
            self, event,
            message="¿Eliminar TODOS los landmarks del caso actual?"
        )


class TESIS_OT_CleanScene(Operator):
    bl_idname = "tesis.clean_scene"
    bl_label = "Limpiar escena"
    bl_description = "Quita objetos visuales de Blender sin borrar los datos guardados en Excel"
    bl_options = {'REGISTER', 'UNDO'}

    def execute(self, context):
        state = context.scene.tesis_state
        if state.dirty:
            self.report({'WARNING'}, "Hay cambios sin guardar. Guarda o descarta los cambios antes de limpiar la escena.")
            return {'CANCELLED'}
        if state.cases and state.active_case < len(state.cases):
            case = state.cases[state.active_case]
            if case.has_neck and state.landmarks:
                statuses = _case_neck_statuses(case, state)
                ccu_incomplete = sum(status == "SET" for status in statuses) == 1 or (
                    len(statuses) == 2 and all(status == "SET" for status in statuses)
                    and _safe_float(case.ccu_mm, 0.0) <= 0
                )
                if ccu_incomplete:
                    self.report({'WARNING'}, "Completa o guarda correctamente la medición cervical antes de limpiar la escena.")
                    return {'CANCELLED'}

        _OBJ_NAME_TO_LM.clear()
        prefixes = ["SKULL_", "SOFT_", "CCU_"] + [f"{i:02d}_" for i in range(len(LANDMARKS_ESPERADOS))]
        removed = 0
        for obj in list(bpy.data.objects):
            if any(obj.name.startswith(prefix) for prefix in prefixes):
                if _safe_remove_object(obj):
                    removed += 1
        for lm in state.landmarks:
            lm.hard_obj_name = ""
            lm.soft_obj_name = ""
            lm.line_obj_name = ""
        invalidate_mallas_cache()
        invalidate_bvh_cache()
        _cleanup_auxiliary_datablocks()
        self.report({'INFO'}, f"Escena limpiada: {removed} objetos visuales eliminados; datos del caso conservados.")
        return {'FINISHED'}

    def invoke(self, context, event):
        return context.window_manager.invoke_confirm(
            self, event,
            message="¿Eliminar TODAS las mallas, esferas, líneas y CCU de la escena?"
        )

class TESIS_OT_NextCase(Operator):
    bl_idname = "tesis.next_case"
    bl_label = "Siguiente caso"
    bl_description = "Avanza al siguiente caso en la lista (con confirmación si hay cambios sin guardar)"

    def execute(self, context):
        state = context.scene.tesis_state

        if not state.cases:
            return {'CANCELLED'}

        _advance_to_next_case(context)
        return {'FINISHED'}

    def invoke(self, context, event):
        state = context.scene.tesis_state

        if state.dirty:
            return context.window_manager.invoke_props_dialog(self)

        return self.execute(context)

    def draw(self, context):
        layout = self.layout
        layout.label(text="Hay cambios sin guardar.")
        layout.label(text="¿Qué deseas hacer?")

        row = layout.row()
        row.operator("tesis.save_and_next", text="Guardar y continuar", icon="CHECKMARK")

        row = layout.row()
        row.operator("tesis.discard_and_next", text="Descartar y continuar", icon="TRASH")

class TESIS_OT_SaveAndNext(Operator):
    bl_idname = "tesis.save_and_next"
    bl_label = "Guardar y avanzar"
    bl_description = "Guarda los landmarks actuales, limpia la escena y carga el siguiente caso pendiente"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases:
            self.report({'WARNING'}, "No hay casos cargados.")
            return {'CANCELLED'}
        case = state.cases[state.active_case]
        blockers = _case_finalize_blockers(state, case)
        if blockers:
            self.report({'WARNING'}, "No se puede finalizar: " + " ".join(blockers))
            return {'CANCELLED'}
        return bpy.ops.tesis.save_case(
            'INVOKE_DEFAULT', finalize_case=True, advance_after_save=True
        )


class TESIS_OT_DiscardAndNext(Operator):
    bl_idname = "tesis.discard_and_next"
    bl_label = "Descartar y avanzar"
    bl_description = "Descarta los cambios actuales sin guardar, limpia la escena y carga el siguiente caso"

    def execute(self, context):
        state = context.scene.tesis_state
        state.dirty = False

        # La transición atómica limpia el caso actual y sincroniza el siguiente.
        if not _advance_to_next_filtered_case(context):
            self.report({'INFO'}, "No hay más casos que cumplan los filtros.")
            return {'FINISHED'}

        # Importar mallas del nuevo caso
        case = state.cases[state.active_case]
        has_paths = (case.ruta_hueso and case.ruta_hueso.strip()
                     and case.ruta_blando and case.ruta_blando.strip())
        if has_paths:
            bpy.ops.tesis.import_meshes('EXEC_DEFAULT')

        return {'FINISHED'}


class TESIS_OT_PrevCase(Operator):
    bl_idname = "tesis.prev_case"
    bl_label = "Caso anterior"
    bl_description = "Retrocede al caso anterior en la lista (con confirmación si hay cambios sin guardar)"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.cases:
            return {'CANCELLED'}
        if state.dirty and bool(getattr(self, "_discard_confirmed", False)):
            state.dirty = False
        if not _advance_to_prev_case(context):
            self.report({'INFO'}, "Ya estás en el primer caso o hay cambios sin guardar.")
        return {'FINISHED'}

    def invoke(self, context, event):
        state = context.scene.tesis_state
        if state.dirty:
            self._discard_confirmed = True
            return context.window_manager.invoke_confirm(
                self, event,
                message="Hay cambios sin guardar. ¿Descartarlos y retroceder?"
            )
        self._discard_confirmed = False
        return self.execute(context)

class TESIS_OT_NextLandmark(bpy.types.Operator):
    bl_idname = "tesis.next_landmark"
    bl_label = "Siguiente Landmark"
    bl_description = "Avanza al siguiente landmark en la lista secuencial (centrales → izq → der → cuello)"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}
        next_idx = _next_sequential_idx(state.active_landmark)
        if next_idx is not None:
            state.active_landmark = next_idx
        return {'FINISHED'}


class TESIS_OT_PrevLandmark(bpy.types.Operator):
    bl_idname = "tesis.prev_landmark"
    bl_label = "Landmark Anterior"
    bl_description = "Retrocede al landmark anterior en la lista secuencial (centrales → izq → der → cuello)"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}
        prev_idx = _prev_sequential_idx(state.active_landmark)
        if prev_idx is not None:
            state.active_landmark = prev_idx
        return {'FINISHED'}


class TESIS_OT_NextMissingLandmark(bpy.types.Operator):
    bl_idname = "tesis.next_missing_landmark"
    bl_label = "Ir al siguiente pendiente"
    bl_description = "Salta automáticamente al próximo landmark pendiente (OMIT) en orden secuencial"

    def execute(self, context):
        state = context.scene.tesis_state
        if not state.landmarks:
            return {'CANCELLED'}

        current_pos = _SEQ_POS_OF.get(state.active_landmark, -1)
        n = len(_SEQUENTIAL_ORDER)

        # buscar adelante en orden secuencial (wrap around)
        for offset in range(1, n):
            pos = (current_pos + offset) % n
            idx = _SEQUENTIAL_ORDER[pos]
            if state.landmarks[idx].status == "OMIT":
                state.active_landmark = idx
                return {'FINISHED'}

        self.report({'INFO'}, "No quedan landmarks pendientes.")
        return {'FINISHED'}


class TESIS_OT_ToggleLandmarkNames(bpy.types.Operator):
    bl_idname = "tesis.toggle_lm_names"
    bl_label = "Mostrar/ocultar nombres de landmarks"
    bl_description = "Activa o desactiva la visualización de etiquetas de texto para los landmarks en la vista 3D"

    def execute(self, context):
        state = context.scene.tesis_state
        state.show_lm_names = not state.show_lm_names

        # aplicar a todos los objetos LM_* (HARD/SOFT)
        for o in bpy.data.objects:
            if o.name.endswith("_HARD") or o.name.endswith("_SOFT"):
                try:
                    o.show_name = state.show_lm_names
                except Exception:
                    pass
        return {'FINISHED'}


class TESIS_OT_ToggleHardVisibility(Operator):
    bl_idname = "tesis.toggle_hard_vis"
    bl_label = "Toggle Hard"
    bl_description = "Mostrar/ocultar esferas de puntos duros (cr\u00e1neo)"

    def execute(self, context):
        col = bpy.data.collections.get("LM_HARD")
        if col:
            col.hide_viewport = not col.hide_viewport
        return {'FINISHED'}


class TESIS_OT_ToggleSoftVisibility(Operator):
    bl_idname = "tesis.toggle_soft_vis"
    bl_label = "Toggle Soft"
    bl_description = "Mostrar/ocultar esferas de puntos blandos"

    def execute(self, context):
        col = bpy.data.collections.get("LM_SOFT")
        if col:
            col.hide_viewport = not col.hide_viewport
        return {'FINISHED'}


class TESIS_OT_ToggleLinesVisibility(Operator):
    bl_idname = "tesis.toggle_lines_vis"
    bl_label = "Toggle L\u00edneas"
    bl_description = "Mostrar/ocultar l\u00edneas de grosor entre puntos"

    def execute(self, context):
        col = bpy.data.collections.get("LM_LINES")
        if col:
            col.hide_viewport = not col.hide_viewport
        return {'FINISHED'}

class TESIS_OT_SelectLandmark(Operator):
    bl_idname = "tesis.select_landmark"
    bl_label = "Seleccionar landmark"
    bl_description = "Selecciona el landmark activo"

    index: IntProperty(default=-1)  # type: ignore

    def execute(self, context):
        state = context.scene.tesis_state
        if self.index < 0 or self.index >= len(state.landmarks):
            return {'CANCELLED'}
        state.active_landmark = self.index
        return {'FINISHED'}


class TESIS_OT_DeleteSelected(Operator):
    bl_idname = "tesis.delete_selected"
    bl_label = "Eliminar selección"
    bl_description = "Elimina el landmark completo seleccionado o usa el borrado normal de Blender"

    def execute(self, context):
        state = context.scene.tesis_state
        obj = getattr(context.view_layer.objects, "active", None)
        entry = _resolve_lm_entry_for_object(state, obj) if obj is not None else None
        if entry is not None:
            idx, _role = entry
            if delete_landmark_complete(state, idx):
                self.report({'INFO'}, "Landmark completo eliminado.")
                return {'FINISHED'}
        return bpy.ops.object.delete('EXEC_DEFAULT')

class TESIS_OT_SyncSelectedLandmark(Operator):
    bl_idname = "tesis.sync_selected_landmark"
    bl_label = "Sincronizar landmark seleccionado"
    bl_description = (
        "Proyecta el landmark seleccionado a su superficie, actualiza coordenadas, "
        "grosor y línea. Es un respaldo manual seguro después de mover con G"
    )

    def execute(self, context):
        state = context.scene.tesis_state
        selected = [obj for obj in context.selected_objects if _resolve_lm_entry_for_object(state, obj)]
        if not selected:
            self.report({'WARNING'}, "Selecciona una esfera HARD o SOFT del landmark.")
            return {'CANCELLED'}
        obj = context.view_layer.objects.active if context.view_layer.objects.active in selected else selected[0]
        ok = _sync_landmark_object_to_surface(
            context.scene, context.evaluated_depsgraph_get(), obj.name, show_header=True
        )
        if not ok:
            self.report({'ERROR'}, "No se pudo sincronizar el landmark seleccionado.")
            return {'CANCELLED'}
        self.report({'INFO'}, "Landmark, superficie, grosor y línea sincronizados.")
        return {'FINISHED'}


class TESIS_OT_SectionInfo(Operator):
    """Muestra instrucciones de cada sección al pasar el mouse por el icono"""
    bl_idname = "tesis.section_info"
    bl_label = "Información"

    section: StringProperty(default="") # type: ignore

    _TOOLTIPS = {
        "CASO": (
            "Seleccione el paciente que desea procesar.\n"
            "\n"
            "Los casos se cargan desde el archivo Excel del proyecto.\n"
            "El Excel contiene información demográfica, rutas de mallas\n"
            "y flags metodológicos.\n"
            "\n"
            "Campo «Ojo» (legacy del Excel):\n"
            "No representa lateralidad ocular. Se usa como flag de alerta\n"
            "para casos que requieren revisión especial por observaciones,\n"
            "patologías, landmarks faltantes o dificultades anatómicas.\n"
            "\n"
            "Campo «Cuello»:\n"
            "Indica si es posible calcular la circunferencia cervical (CCU).\n"
            "1 = se puede calcular CCU | 0 = no se puede calcular CCU"
        ),
        "IMPORT": (
            "Importe las mallas correspondientes al paciente seleccionado.\n"
            "\n"
            "El sistema utiliza dos superficies:\n"
            "• SKULL → superficie ósea (cráneo)\n"
            "• SOFT → superficie cutánea (tejido blando)\n"
            "\n"
            "Las mallas deben estar alineadas y en escala correcta.\n"
            "\n"
            "Los landmarks se marcan en el cráneo y luego se proyectan\n"
            "sobre el tejido blando mediante ray casting y BVH trees\n"
            "para calcular el grosor de tejido blando"
        ),
        "LANDMARKS": (
            "Marque los landmarks craneofaciales sobre la superficie del cráneo.\n"
            "\n"
            "El flujo de marcado sigue el orden anatómico definido\n"
            "en LANDMARKS_ESPERADOS.\n"
            "\n"
            "Para cada landmark:\n"
            "1. Seleccione el punto en el cráneo\n"
            "2. El sistema proyecta automáticamente al tejido blando\n"
            "3. Se calcula el grosor entre ambas superficies\n"
            "\n"
            "Estados posibles:\n"
            "• SET → landmark marcado\n"
            "• OMIT → pendiente de marcar\n"
            "• MISSING → no aplicable anatómicamente"
        ),
        "ANALISIS": (
            "Verifique las mediciones obtenidas.\n"
            "\n"
            "El grosor del tejido blando se calcula como la distancia entre:\n"
            "• landmark HARD (superficie ósea)\n"
            "• landmark SOFT (superficie cutánea)\n"
            "\n"
            "Utilizando intersección geométrica con BVH trees.\n"
            "\n"
            "Use el diagnóstico de coordenadas para revisar casos incompatibles"
        ),
        "CUELLO": (
            "Calcule la circunferencia cervical (CCU).\n"
            "\n"
            "Solo se utiliza en casos donde el campo «Cuello» está activado.\n"
            "\n"
            "Landmarks requeridos:\n"
            "• neck_ant → punto anatómico anterior del cuello\n"
            "• neck_post → punto anatómico posterior del cuello\n"
            "\n"
            "El addon generará un anillo alrededor del cuello utilizando\n"
            "estos puntos como referencia geométrica para estimar la CCU.\n"
            "\n"
            "La medición se usa en este proyecto como variable exploratoria\n"
            "asociada al estado nutricional del individuo"
        ),
        "GUARDAR": (
            "Guarde los resultados del caso actual.\n"
            "\n"
            "Se exportarán al archivo Excel del proyecto:\n"
            "• Coordenadas de landmarks (hoja Coordenadas)\n"
            "• Grosor de tejido blando (hoja Resultados)\n"
            "• Medición CCU y estado nutricional\n"
            "\n"
            "El Excel funciona como base de datos central del proyecto.\n"
            "«Guardar y Siguiente» limpia la escena y carga el próximo caso"
        ),
        "HERRAMIENTAS": (
            "Herramientas del flujo activo del caso.\n"
            "\n"
            "• Actualizar BVH: reconstruye la caché de intersección para la malla activa.\n"
            "• Reconstruir LM: rehidrata landmarks previamente guardados desde Excel.\n"
            "• Importar mallas: carga SKULL y SOFT del paciente actual.\n"
            "• Toggle Hard/Soft/Líneas: alterna la visibilidad de los elementos del caso.\n"
            "• Limpiar escena: elimina solo los elementos visuales del caso para trabajar limpio.\n"
            "\n"
            "Estas opciones apoyan la revisión y el reflujo del caso activo sin alterar la geometría base."
        ),
    }

    @classmethod
    def description(cls, context, properties):
        return cls._TOOLTIPS.get(properties.section, "Información de la sección")

    def execute(self, context):
        return {'FINISHED'}


# ─────────────────────────────────────────────────────
# PANELES — Workflow estructurado por etapas
# ─────────────────────────────────────────────────────

def _draw_section_help_icon(layout, section: str):
    op = layout.operator("tesis.section_info", text="", icon='QUESTION', emboss=False)
    op.section = section


class TESIS_PT_MainPanel(Panel):
    bl_label = "TESIS Landmark System"
    bl_idname = "TESIS_PT_panel"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_options = set()  # Panel principal siempre abierto, sin colapsar

    def draw(self, context):
        layout = self.layout
        layout.use_property_split = False
        state = context.scene.tesis_state

        if not state.cases:
            layout.label(text="TESIS Landmark System", icon='BONE_DATA')
            draw_wrapped_paragraphs(
                layout,
                "Antropología Física · Universidad de Chile\n"
                "Josefa Diéguez Burgos\n\n"
                "Sistema de apoyo para el registro y medición de landmarks craneofaciales y espesores de tejido blando.\n\n"
                "1. Carga el Excel maestro del proyecto.\n"
                "2. Selecciona un caso.\n"
                "3. Importa las mallas para comenzar.",
                context,
            )
            layout.operator("tesis.load_excel", text="Cargar Excel maestro", icon='FILE_FOLDER')
            return

        total_cases = state.case_count_total or len(state.cases)
        processed_cases = state.case_count_processed
        pending_cases = state.case_count_pending if state.case_count_total else max(0, total_cases - processed_cases)
        excel_path = get_excel_path()
        excel_name = os.path.basename(excel_path) if excel_path else "No configurado"

        # ── Info del proyecto ────────────────────────────────────
        info = layout.row(align=True)
        info.scale_y = 0.82
        info.label(text=excel_name, icon='FILE_CACHE')
        if total_cases:
            info.label(text=f"{processed_cases}/{total_cases} casos  \u00b7  {pending_cases} pendientes")

        # ── Flujo principal (arranque) ───────────────────────────
        col = layout.column(align=True)
        col.scale_y = 1.1
        r1 = col.row(align=True)
        r1.operator("tesis.load_excel", text="Recargar Excel" if state.cases else "Cargar Excel maestro", icon='FILE_FOLDER')

        if state.active_case >= len(state.cases):
            return

        case = state.cases[state.active_case]
        skull, soft = obtener_mallas_paciente(quiet=True)
        meshes_ok = skull is not None and soft is not None
        lm_total = len(_OTHER_INDICES)

        layout.separator(factor=0.4)

        # El caso activo se detalla únicamente en Caso y filtros.
        if not case.coordinate_frame_valid or case.coordinate_review_required:
            warning = layout.row()
            warning.alert = True
            warning.label(text="Coordenadas históricas requieren revisión", icon='ERROR')


# ── 1. CASO ─────────────────────────────────────────────


class TESIS_PT_CasoPanel(Panel):
    bl_label = "1. Caso y filtros"
    bl_idname = "TESIS_PT_caso"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        return bool(context.scene.tesis_state.cases)

    def draw_header(self, context):
        state = context.scene.tesis_state
        if state.cases and state.active_case < len(state.cases):
            self.layout.label(text="", icon='CHECKMARK')
        _draw_section_help_icon(self.layout, "CASO")

    def draw(self, context):
        layout = self.layout
        state = context.scene.tesis_state
        if not state.cases:
            layout.label(text="Primero carga casos desde Excel.", icon='INFO')
            return
        if state.active_case >= len(state.cases):
            return

        case = state.cases[state.active_case]
        nav = layout.row(align=True)
        nav.operator("tesis.prev_case", text="", icon='TRIA_LEFT')
        nav.label(text=case.paciente, icon='USER')
        nav.operator("tesis.next_case", text="", icon='TRIA_RIGHT')

        # Contexto de medición: principal vs. cohorte de reproducibilidad.
        measurement = layout.box()
        measurement.label(text="Contexto de medición", icon='GROUP')
        measurement.prop(state, "measurement_mode", text="Modo")
        if state.measurement_mode == "ERROR":
            alert = measurement.row()
            alert.alert = True
            alert.label(text="ESTUDIO DE ERROR · no sobrescribe la medición principal", icon='INFO')
            measurement.prop(state, "error_cohort", text="Cohorte")
            row = measurement.row(align=True)
            row.prop(state, "observer_id", text="Observador")
            row.prop(state, "session_id", text="Sesión")
            row = measurement.row(align=True)
            row.prop(state, "error_only_pending", text="Solo pendientes")
            row.prop(state, "error_include_neck", text="Incluir CCU")

            total_err, done_err, pending_err = _error_session_counts(state)
            measurement.label(
                text=f"Sesión actual: {done_err}/{total_err} completos · {pending_err} pendientes",
                icon='CHECKMARK' if total_err and pending_err == 0 else 'TIME',
            )
            existing = any(
                bool(getattr(c, "error_study", False))
                and _safe_str(getattr(c, "error_cohort", "")) == _safe_str(state.error_cohort)
                for c in state.cases
            )
            if not existing:
                setup = measurement.column(align=True)
                setup.label(text="Crear cohorte una sola vez")
                row = setup.row(align=True)
                row.prop(state, "error_sample_size", text="N")
                row.prop(state, "error_seed", text="Semilla")
                setup.operator("tesis.generate_error_cohort", text="Generar cohorte estratificada", icon='GROUP')
            else:
                measurement.operator("tesis.generate_error_cohort", text="Usar cohorte guardada", icon='GROUP')
        else:
            row = measurement.row(align=True)
            row.prop(state, "observer_id", text="Observador")
            row.prop(state, "session_id", text="Sesión")
            row.prop(state, "replicate_id", text="Réplica")
            if _safe_str(state.session_id).upper() != _PRIMARY_MEASUREMENT_SESSION:
                warn = measurement.row()
                warn.alert = True
                warn.label(text=f"Principal canónico = {_PRIMARY_MEASUREMENT_SESSION}; no se guardará otra sesión como NORMAL", icon='ERROR')
            else:
                measurement.label(text="Principal canónico · NORMAL / S1", icon='CHECKMARK')

        filters = layout.box()
        top = filters.row(align=True)
        visible_count = sum(1 for c in state.cases if _case_matches_active_filters(state, c))
        top.label(text=f"Filtros · {visible_count}/{len(state.cases)} visibles", icon='FILTER')
        top.operator("tesis.reset_filters", text="", icon='X')

        filters.prop(state, "search", text="Buscar", icon='VIEWZOOM')
        r = filters.row(align=True)
        r.prop(state, "case_status", text="Estado")
        r.prop(state, "filter_sex", text="Sexo")
        r = filters.row(align=True)
        r.prop(state, "filter_age_band", text="Grupo")
        r.prop(state, "filter_neck", text="Cuello")
        r = filters.row(align=True)
        r.prop(state, "filter_has_obs", text="Observaciones")
        r = filters.row(align=True)
        r.prop(state, "filter_age_min", text="Edad mín.")
        r.prop(state, "filter_age_max", text="Edad máx.")
        if state.dirty:
            notice = filters.row()
            notice.alert = True
            notice.label(text="Hay cambios sin guardar. Guarda o descarta antes de cambiar de caso.", icon='ERROR')

        box = layout.box()
        head = box.row(align=True)
        head.scale_y = 0.75
        _draw_case_columns(head, ["Paciente", "Edad", "Sexo", "Grupo", "LM", "Estado"])
        box.template_list("TESIS_UL_Cases", "", state, "cases", state, "active_case", rows=6)

        details = layout.box()
        h = details.row(align=True)
        h.label(text="Detalles del caso", icon='INFO')
        details.label(text=f"Paciente {case.paciente} · {case.edad}a · {case.sexo} · {case.grupo}")
        details.label(text=f"{state.lm_count_set} medidos · {state.lm_count_missing} N/D · {state.lm_count_pending} pendientes")
        if state.measurement_mode == "ERROR":
            session_complete = _error_session_complete(case, state)
            details.label(
                text=f"Cohorte {case.error_cohort or '—'} · orden {case.error_order or '—'} · "
                     f"{state.observer_id}/{state.session_id}/R{state.replicate_id}",
                icon='GROUP',
            )
            details.label(text=f"Sesión: {'COMPLETA' if session_complete else 'EN PROGRESO'}")
        else:
            details.label(text=f"Estado: {'COMPLETO' if _case_is_complete(case, state=state) else 'EN PROGRESO'}")
        if state.measurement_mode == "ERROR":
            details.label(text=f"Observaciones: {_safe_str(case.obs) or '—'}")
            row = details.row(align=True)
            row.label(text=f"Tiene cuello: {'Sí' if case.has_neck else 'No'}")
        else:
            details.prop(case, "obs", text="Observaciones")
            row = details.row(align=True)
            row.label(text="Tiene cuello")
            toggle = row.operator("tesis.toggle_case_neck", text="Sí" if case.has_neck else "No", icon='CHECKBOX_HLT' if case.has_neck else 'CHECKBOX_DEHLT')
            toggle.value = not bool(case.has_neck)
        if case.has_neck:
            neck_statuses = _case_neck_statuses(case, state)
            ant = "✓" if len(neck_statuses) > 0 and neck_statuses[0] == "SET" else "—"
            post = "✓" if len(neck_statuses) > 1 and neck_statuses[1] == "SET" else "—"
            ccu = f"CCU {case.ccu_mm:.1f}" if case.ccu_mm > 0 else "CCU pendiente"
            details.label(text=f"Cuello: Anterior {ant} / Posterior {post} / {ccu}")


# ── 2. IMPORTACIÓN ─────────────────────────────────────

class TESIS_PT_ImportPanel(Panel):
    bl_label = "2. Importación"
    bl_idname = "TESIS_PT_import"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        return bool(context.scene.tesis_state.cases)

    def draw_header(self, context):
        skull, soft = obtener_mallas_paciente(quiet=True)
        if skull and soft:
            self.layout.label(text="", icon='CHECKMARK')
        _draw_section_help_icon(self.layout, "IMPORT")

    def draw(self, context):
        layout = self.layout
        skull, soft = obtener_mallas_paciente(quiet=True)
        meshes_ok = skull is not None and soft is not None

        row = layout.row(align=True)
        row.scale_y = 1.1
        busy = TESIS_OT_ImportMeshes._busy
        row.enabled = not busy
        row.operator("tesis.import_meshes", text="Importando…" if busy else "Importar caso", icon='IMPORT')
        tip = layout.box()
        tip.scale_y = 0.85
        tip.label(
            text="Mallas listas para marcar." if meshes_ok else "Importar reconstruye landmarks guardados",
            icon='CHECKMARK' if meshes_ok else 'INFO'
        )


class TESIS_PT_LandmarksPanel(Panel):
    bl_label = "3. Landmarks"
    bl_idname = "TESIS_PT_landmarks"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        state = context.scene.tesis_state
        return bool(state.cases and state.landmarks)

    def draw_header(self, context):
        state = context.scene.tesis_state
        if state.landmarks:
            icon = 'CHECKMARK' if state.lm_count_pending == 0 else 'BONE_DATA'
            self.layout.label(text="", icon=icon)
        _draw_section_help_icon(self.layout, "LANDMARKS")

    def draw(self, context):
        layout = self.layout
        state = context.scene.tesis_state

        if not state.landmarks:
            layout.label(text="Importa las mallas del caso para generar landmarks.", icon='INFO')
            return

        skull, soft = obtener_mallas_paciente(quiet=True)
        meshes_ok = skull is not None and soft is not None
        if not meshes_ok:
            notice = layout.row()
            notice.label(text="Modo lectura · importa las mallas para marcar o corregir.", icon='INFO')

        _draw_quick_tools(layout, context)
        layout.separator(factor=0.35)

        # ── Navegación y marcado secuencial ──────────────────────
        # IMPORTANTE: esta acción usa el mismo operador funcional de siempre
        # (TESIS_OT_AddLandmark con sequential=True). El rediseño del panel no
        # debe convertir el flujo continuo en 26 marcaciones aisladas.
        if meshes_ok:
            nav = layout.row(align=True)
            nav.operator("tesis.prev_landmark", text="", icon='TRIA_LEFT')
            idx = state.active_landmark
            if 0 <= idx < len(LANDMARKS_ESPERADOS):
                start = nav.operator("tesis.add_landmark", text=_ui_landmark_label(idx), icon='PLAY')
                start.sequential = True
            nav.operator("tesis.next_landmark", text="", icon='TRIA_RIGHT')
            nav.operator("tesis.next_missing_landmark", text="", icon='VIEWZOOM')

            seq = layout.row(align=True)
            seq.scale_y = 1.05
            start = seq.operator("tesis.add_landmark", text="Iniciar / continuar marcado secuencial", icon='PLAY')
            start.sequential = True

        # Estado del landmark activo: un único control de disponibilidad.
        availability = layout.row(align=True)
        availability.scale_y = 0.90
        idx = state.active_landmark
        if 0 <= idx < len(state.landmarks) and idx in _OTHER_INDICES:
            lm_active = state.landmarks[idx]
            availability.label(text=f"Activo: {_ui_landmark_label(idx)}")
            button_text = "Volver a pendiente" if lm_active.status == "MISSING" else "No disponible"
            button_icon = 'LOOP_BACK' if lm_active.status == "MISSING" else 'CANCEL'
            op = availability.operator(
                "tesis.set_landmark_status",
                text=button_text,
                icon=button_icon,
            )
            op.index = -1
            op.status = "MISSING"
        else:
            availability.enabled = False
            availability.label(text="Disponibilidad facial")

        # ── Lista compacta de landmarks ───────────────────────────
        _draw_landmark_grid(layout.column(align=True), state, meshes_ok=meshes_ok)


# ── 4. CUELLO (CCU) ────────────────────────────────────

class TESIS_PT_CuelloPanel(Panel):
    bl_label = "4. Cuello (CCU)"
    bl_idname = "TESIS_PT_cuello"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        return bool(context.scene.tesis_state.cases)

    def draw_header(self, context):
        state = context.scene.tesis_state
        case = state.cases[state.active_case] if state.cases and state.active_case < len(state.cases) else None
        if case and case.has_neck and case.ccu_mm > 0:
            self.layout.label(text="", icon='CHECKMARK')
        _draw_section_help_icon(self.layout, "CUELLO")

    def draw(self, context):
        layout = self.layout
        state = context.scene.tesis_state
        case = state.cases[state.active_case] if state.cases else None
        if case is None:
            return
        if not case.has_neck:
            layout.label(text="Cuello no disponible para este caso", icon='INFO')
            return

        lm_ant = find_landmark_by_name(state.landmarks, "neck_ant") if state.landmarks else None
        lm_post = find_landmark_by_name(state.landmarks, "neck_post") if state.landmarks else None

        skull, soft = obtener_mallas_paciente(quiet=True)
        meshes_ok = skull is not None and soft is not None

        # Los dos puntos cervicales se marcan de forma totalmente independiente.
        # No existe avance automático Anterior -> Posterior.
        points = layout.row(align=True)
        points.scale_y = 1.0

        ant = points.row(align=True)
        ant.enabled = meshes_ok
        ant_op = ant.operator(
            "tesis.mark_neck_point",
            text="Anterior",
            icon='CHECKMARK' if lm_ant and lm_ant.status == "SET" else 'CHECKBOX_DEHLT',
        )
        ant_op.target = "neck_ant"

        post = points.row(align=True)
        post.enabled = meshes_ok
        post_op = post.operator(
            "tesis.mark_neck_point",
            text="Posterior",
            icon='CHECKMARK' if lm_post and lm_post.status == "SET" else 'CHECKBOX_DEHLT',
        )
        post_op.target = "neck_post"

        ready = bool(lm_ant and lm_post and lm_ant.status == "SET" and lm_post.status == "SET")
        ring = _get_current_ccu_ring(case, soft) if soft is not None else None

        actions = layout.row(align=True)
        create = actions.row(align=True)
        create.enabled = ready and meshes_ok
        create.operator("tesis.create_ccu_ring", text="Crear / regenerar anillo", icon='MESH_CIRCLE')
        confirm = actions.row(align=True)
        confirm.enabled = ring is not None
        confirm.operator("tesis.confirm_ccu", text="Confirmar CCU", icon='CHECKMARK')

        if case.ccu_mm > 0:
            layout.label(text=f"CCU guardada: {case.ccu_mm:.2f}", icon='FILE_TICK')
            if case.estado_nutricional:
                layout.label(text=case.estado_nutricional, icon='INFO')

        if ring is not None:
            visual = float(ring.get("ccu_visual_length_mm", 0.0))
            if visual <= 0:
                visual = _compute_ring_length_mm(ring)
            if visual > 0:
                layout.label(text=f"CCU visual: {visual:.2f}", icon='MESH_CIRCLE')

            estimated = int(ring.get("ccu_interpolated_points", 0))
            if estimated > 0:
                q = layout.row()
                q.alert = True
                q.label(text=f"Contorno estimado: {estimated}/256 puntos", icon='ERROR')
            else:
                layout.label(text="Contorno directo sobre piel", icon='CHECKMARK')
        elif case.ccu_mm > 0 and ready:
            pending = layout.row()
            pending.alert = True
            pending.label(text="CCU guardada · anillo visual pendiente de regenerar", icon='ERROR')


class TESIS_PT_GuardarPanel(Panel):
    bl_label = "5. Guardar"
    bl_idname = "TESIS_PT_guardar"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        return bool(context.scene.tesis_state.cases)

    def draw_header(self, context):
        state = context.scene.tesis_state
        self.layout.label(text="", icon='FILE_TICK' if not state.dirty else 'ERROR')
        _draw_section_help_icon(self.layout, "GUARDAR")

    def draw(self, context):
        layout = self.layout
        state = context.scene.tesis_state
        case = state.cases[state.active_case] if state.cases and state.active_case < len(state.cases) else None

        if case is not None:
            warnings = []

            if state.dirty:
                warnings.append("sin guardar")
            if state.lm_count_pending > 0:
                warnings.append(f"{state.lm_count_pending} pendientes")
            if case.has_neck and case.ccu_mm <= 0:
                warnings.append("CCU pendiente")
            if case.coordinate_review_required:
                warnings.append("revisión de coordenadas")

            if warnings:
                row = layout.row(align=True)
                row.alert = True
                row.label(text=" | ".join(warnings), icon='ERROR')

        if state.measurement_mode == "ERROR":
            mode_box = layout.box()
            mode_box.label(
                text=f"ERROR · {state.error_cohort} · {state.observer_id}/{state.session_id}/R{state.replicate_id}",
                icon='GROUP',
            )
            mode_box.label(text="Se guarda solo en Mediciones_LM; la medición principal queda intacta.", icon='INFO')

        row = layout.row(align=True)
        row.scale_y = 1.05

        saving = _SAVE_IN_PROGRESS

        save_row = row.row(align=True)
        frame_valid = (state.cases
                   and state.cases[state.active_case].coordinate_frame_valid
                   and not state.cases[state.active_case].coordinate_review_required)
        save_row.enabled = state.dirty and not saving and frame_valid
        save_label = "Guardar medición" if state.measurement_mode == "ERROR" else "Guardar progreso"
        op = save_row.operator("tesis.save_case", text="Guardando…" if saving else save_label, icon='FILE_TICK')
        op.finalize_case = False
        op.advance_after_save = False

        next_row = row.row(align=True)
        blockers = _case_finalize_blockers(state, case) if state.cases else ["No hay caso activo."]
        next_row.enabled = not saving and frame_valid and not blockers
        next_row.operator("tesis.save_and_next", text="Guardar y siguiente  Shift↵", icon='FORWARD')
        if blockers:
            reason = layout.row()
            reason.alert = True
            reason.label(text="No se puede finalizar: " + blockers[0], icon='ERROR')


# ── HERRAMIENTAS ────────────────────────────────────────

class TESIS_PT_HerramientasPanel(Panel):
    bl_label = "Herramientas"
    bl_idname = "TESIS_PT_herramientas"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "TESIS"
    bl_parent_id = "TESIS_PT_panel"

    @classmethod
    def poll(cls, context):
        return bool(context.scene.tesis_state.cases)

    def draw_header(self, context):
        _draw_section_help_icon(self.layout, "HERRAMIENTAS")

    def draw(self, context):
        layout = self.layout
        state = context.scene.tesis_state

        # Las acciones frecuentes de visibilidad y BVH viven únicamente en la
        # barra compacta de Landmarks. La reconstrucción sigue disponible por F3.
        if state.cases and state.active_case < len(state.cases):
            case = state.cases[state.active_case]
            if case.coordinate_review_required:
                warn = layout.box()
                warn.alert = True
                warn.label(text="Coordenadas históricas no alineadas", icon='ERROR')
                warn.operator(
                    "tesis.diagnose_coordinate_frame",
                    text="Diagnosticar coordenadas",
                    icon='VIEWZOOM'
                )

        maintenance = layout.box()
        maintenance.label(text="Acciones de mantenimiento", icon='TOOL_SETTINGS')
        row = maintenance.row(align=True)
        row.alert = True
        row.operator("tesis.clear_current_case", text="Borrar LM", icon='TRASH')
        row.alert = False
        row.operator("tesis.clean_scene", text="Limpiar escena", icon='X')


def _sync_landmark_state_from_evaluated(scene, depsgraph, *, mark_dirty=True) -> int:
    """Lee las posiciones finales de los objetos y actualiza el estado del addon.

    No modifica mallas, materiales ni constraints. Se usa en sincronizaciones
    explícitas como guardado o reparación de referencias.
    """
    if _undo_in_progress:
        return 0
    state = getattr(scene, "tesis_state", None)
    if state is None or not getattr(state, "landmarks", None):
        return 0

    changed = 0
    for idx, lm in enumerate(state.landmarks):
        if lm.status != "SET":
            continue
        soft_obj = bpy.data.objects.get(getattr(lm, "soft_obj_name", ""))
        if soft_obj is None:
            continue
        p_soft = _evaluated_world_location(soft_obj, depsgraph)
        if p_soft is None:
            continue

        is_neck = normalize_landmark_name(lm.name) in _NECK_IDS
        old_soft = Vector((lm.x, lm.y, lm.z))
        position_changed = (p_soft - old_soft).length > 1e-5
        lm.x, lm.y, lm.z = p_soft.x, p_soft.y, p_soft.z

        if is_neck:
            lm.thickness = 0.0
        else:
            hard_obj = bpy.data.objects.get(getattr(lm, "hard_obj_name", ""))
            p_hard = _evaluated_world_location(hard_obj, depsgraph) if hard_obj else Vector((lm.hx, lm.hy, lm.hz))
            if p_hard is None:
                continue
            old_hard = Vector((lm.hx, lm.hy, lm.hz))
            old_th = float(lm.thickness)
            lm.hx, lm.hy, lm.hz = p_hard.x, p_hard.y, p_hard.z
            lm.thickness = (p_soft - p_hard).length
            position_changed = (
                position_changed
                or (p_hard - old_hard).length > 1e-5
                or abs(float(lm.thickness) - old_th) > 1e-5
            )

            line_obj = bpy.data.objects.get(getattr(lm, "line_obj_name", ""))
            if line_obj is not None:
                # Actualización explícita y segura desde timer. Se ejecuta aunque
                # una versión anterior haya dejado drivers o flags de vínculo.
                if position_changed:
                    update_line_endpoints(line_obj, p_hard, p_soft)

        if position_changed:
            changed += 1

    if changed and mark_dirty:
        state.dirty = True
        try:
            for area in bpy.context.screen.areas:
                if area.type == 'VIEW_3D':
                    area.tag_redraw()
        except Exception:
            pass
    return changed


class TESIS_OT_RepairLandmarkLinks(Operator):
    bl_idname = "tesis.repair_landmark_links"
    bl_label = "Reparar vínculos LM"
    bl_description = "Restaura adherencia a superficie, línea dinámica y referencias de los landmarks existentes"

    def execute(self, context):
        state = context.scene.tesis_state
        skull, soft_mesh = obtener_mallas_paciente(quiet=True)
        if soft_mesh is None:
            self.report({'WARNING'}, "Importa primero las mallas del caso.")
            return {'CANCELLED'}

        repaired = 0
        for idx, lm in enumerate(state.landmarks):
            if lm.status != "SET":
                continue
            soft_obj = bpy.data.objects.get(getattr(lm, "soft_obj_name", ""))
            hard_obj = bpy.data.objects.get(getattr(lm, "hard_obj_name", ""))
            if soft_obj is None:
                continue
            _register_lm_objs(idx, soft_obj.name, hard_obj.name if hard_obj else "")
            _remove_soft_surface_constraint(soft_obj)

            if normalize_landmark_name(lm.name) not in _NECK_IDS and hard_obj is not None:
                line_obj = bpy.data.objects.get(getattr(lm, "line_obj_name", ""))
                if line_obj is None:
                    depsgraph = context.evaluated_depsgraph_get()
                    p_hard = _evaluated_world_location(hard_obj, depsgraph)
                    p_soft = _evaluated_world_location(soft_obj, depsgraph)
                    if p_hard is not None and p_soft is not None:
                        line_obj = create_line_object(
                            f"{idx:02d}_{lm.name}_LINE", p_hard, p_soft,
                            hard_obj=hard_obj, soft_obj=soft_obj,
                        )
                        lm.line_obj_name = line_obj.name
                if line_obj is not None:
                    # Reintentar el vínculo dinámico; si falla, dejar fallback manual.
                    try:
                        spline = line_obj.data.splines[0]
                        for point_idx in range(min(2, len(spline.points))):
                            point = spline.points[point_idx]
                            for axis in range(3):
                                _remove_curve_point_driver(point, axis)
                    except Exception:
                        pass
                    depsgraph = context.evaluated_depsgraph_get()
                    p_hard = _evaluated_world_location(hard_obj, depsgraph)
                    p_soft = _evaluated_world_location(soft_obj, depsgraph)
                    if p_hard is not None and p_soft is not None:
                        line_bound = _bind_line_to_landmark_objects(line_obj, hard_obj, soft_obj)
                        update_line_endpoints(line_obj, p_hard, p_soft)
                        try:
                            line_obj["tesis_line_bound"] = bool(line_bound)
                        except Exception:
                            pass
            repaired += 1

        context.view_layer.update()
        _sync_landmark_state_from_evaluated(
            context.scene, context.evaluated_depsgraph_get(), mark_dirty=False
        )
        self.report({'INFO'}, f"Vínculos reparados: {repaired} landmarks.")
        return {'FINISHED'}


def _tag_tesis_ui_redraw(context):
    """Redibuja solamente las VIEW_3D necesarias para reflejar datos LM."""
    redrawn_areas = set()
    try:
        if context.area is not None and context.area.type == 'VIEW_3D':
            context.area.tag_redraw()
            redrawn_areas.add(context.area.as_pointer())
            for region in context.area.regions:
                region.tag_redraw()
    except Exception:
        pass

    try:
        screen = context.screen
        if screen is not None:
            for area in screen.areas:
                if area.type == 'VIEW_3D' and area.as_pointer() not in redrawn_areas:
                    area.tag_redraw()
                    for region in area.regions:
                        region.tag_redraw()
    except Exception:
        pass


class TESIS_OT_MoveLandmarkSurface(Operator):
    """Mueve un landmark mediante el primer impacto visible del viewport."""
    bl_idname = "tesis.move_landmark_surface"
    bl_label = "Mover landmark sobre superficie"
    bl_description = "Mueve el landmark sobre la primera superficie visible"
    bl_options = {'REGISTER', 'BLOCKING'}

    def _update_from_hit(self, context, hit_point):
        state = context.scene.tesis_state
        if self.idx < 0 or self.idx >= len(state.landmarks):
            return False
        obj = bpy.data.objects.get(self.obj_name)
        if obj is None:
            return False
        lm = state.landmarks[self.idx]
        point = Vector(hit_point)
        if getattr(self, "_last_hit_point", None) is not None and (point - self._last_hit_point).length <= 1e-4:
            return False
        self._last_hit_point = point.copy()
        _set_object_world_location(obj, point)

        if self.is_neck:
            lm.x, lm.y, lm.z = point.x, point.y, point.z
            lm.thickness = 0.0
        else:
            if self.role == "SOFT":
                p_soft = point
                p_hard = Vector((lm.hx, lm.hy, lm.hz))
                lm.x, lm.y, lm.z = p_soft.x, p_soft.y, p_soft.z
            elif self.role == "HARD":
                p_hard = point
                p_soft = Vector((lm.x, lm.y, lm.z))
                lm.hx, lm.hy, lm.hz = p_hard.x, p_hard.y, p_hard.z
            else:
                return False

            lm.thickness = (p_soft - p_hard).length
            line_obj = bpy.data.objects.get(lm.line_obj_name)
            if line_obj is not None:
                update_line_endpoints(line_obj, p_hard, p_soft)
        state.active_landmark = self.idx
        state.dirty = True
        self._moved_valid = True

        try:
            context.scene.update_tag()
        except Exception:
            pass

        now = time.monotonic()
        if now - self._last_ui_redraw >= 1.0 / 30.0:
            _tag_tesis_ui_redraw(context)
            self._last_ui_redraw = now
        return True

    def _restore(self, context):
        state = context.scene.tesis_state
        obj = bpy.data.objects.get(self.obj_name)
        if obj is not None:
            _set_object_world_location(obj, self.original_position)
        if self.idx < len(state.landmarks):
            lm = state.landmarks[self.idx]
            for name, value in self.original_values.items():
                setattr(lm, name, value)
            hard_obj = bpy.data.objects.get(lm.hard_obj_name)
            soft_obj = bpy.data.objects.get(lm.soft_obj_name)
            line_obj = bpy.data.objects.get(lm.line_obj_name)
            if hard_obj is not None and soft_obj is not None and line_obj is not None:
                update_line_endpoints(
                    line_obj,
                    hard_obj.matrix_world.translation,
                    soft_obj.matrix_world.translation,
                )

    def invoke(self, context, event):
        obj = getattr(context.view_layer.objects, 'active', None)
        state = context.scene.tesis_state
        if obj is None or not obj.select_get():
            return bpy.ops.transform.translate('INVOKE_DEFAULT')
        entry = _resolve_lm_entry_for_object(state, obj)
        if entry is None:
            return bpy.ops.transform.translate('INVOKE_DEFAULT')
        if context.area is None or context.area.type != 'VIEW_3D':
            return bpy.ops.transform.translate('INVOKE_DEFAULT')

        skull, soft = obtener_mallas_paciente(quiet=True)
        if skull is None or soft is None:
            return {'CANCELLED'}
        ensure_bvh_cache(context, skull, soft)
        self.obj_name = obj.name
        self.idx, self.role = entry
        self.is_neck = normalize_landmark_name(state.landmarks[self.idx].name) in _NECK_IDS
        if self.role == "HARD":
            self.bvh = _BVH_CACHE["skull"]
            self.fallback_bvh = None
        else:
            self.bvh = get_full_soft_bvh()
            full_soft_bvh = _BVH_CACHE.get("soft")
            self.fallback_bvh = (
                full_soft_bvh
                if self.is_neck and full_soft_bvh is not None and full_soft_bvh is not self.bvh
                else None
            )
        self.region = next((r for r in context.area.regions if r.type == 'WINDOW'), None)
        self.rv3d = context.space_data.region_3d
        self.original_position = obj.matrix_world.translation.copy()
        self.original_dirty = bool(state.dirty)
        self._last_ui_redraw = 0.0
        self._next_move_time = 0.0
        self._last_hit_point = None
        self._moved_valid = False
        lm = state.landmarks[self.idx]
        self.original_values = {
            name: float(getattr(lm, name))
            for name in ("x", "y", "z", "hx", "hy", "hz", "thickness")
        }
        context.window_manager.modal_handler_add(self)
        return {'RUNNING_MODAL'}

    def modal(self, context, event):
        if event.type == 'MOUSEMOVE' and self.region is not None:
            now = time.monotonic()
            if now < self._next_move_time:
                return {'RUNNING_MODAL'}
            self._next_move_time = now + 1.0 / 60.0
            coord_x = event.mouse_x - self.region.x
            coord_y = event.mouse_y - self.region.y
            if not (0 <= coord_x < self.region.width and 0 <= coord_y < self.region.height):
                return {'RUNNING_MODAL'}
            coord = (coord_x, coord_y)
            origin = view3d_utils.region_2d_to_origin_3d(self.region, self.rv3d, coord)
            direction = view3d_utils.region_2d_to_vector_3d(self.region, self.rv3d, coord).normalized()
            hit = self.bvh.ray_cast(origin, direction)[0] if self.bvh is not None else None
            if hit is None and self.fallback_bvh is not None:
                hit = self.fallback_bvh.ray_cast(origin, direction)[0]
            if hit is not None:
                self._update_from_hit(context, hit)
            return {'RUNNING_MODAL'}
        if event.type in {'LEFTMOUSE', 'RET', 'NUMPAD_ENTER'} and event.value == 'PRESS':
            if self.is_neck and self._moved_valid:
                state = context.scene.tesis_state
                case = state.cases[state.active_case] if state.cases else None
                _, soft_obj = obtener_mallas_paciente(quiet=True)
                _remove_current_ccu_ring(case, soft_obj)
                if case is not None and float(getattr(case, "ccu_mm", 0.0)) > 0:
                    self.report(
                        {'WARNING'},
                        "Punto cervical movido. La CCU guardada no cambió; regenere y confirme el anillo."
                    )
            _tag_tesis_ui_redraw(context)
            return {'FINISHED'}
        if event.type in {'ESC', 'RIGHTMOUSE'} and event.value == 'PRESS':
            self._restore(context)
            context.scene.tesis_state.dirty = self.original_dirty
            _tag_tesis_ui_redraw(context)
            return {'CANCELLED'}
        if event.type in {'MIDDLEMOUSE', 'WHEELUPMOUSE', 'WHEELDOWNMOUSE', 'TRACKPADPAN', 'TRACKPADZOOM'}:
            return {'PASS_THROUGH'}
        return {'RUNNING_MODAL'}

    def cancel(self, context):
        self._restore(context)
        _tag_tesis_ui_redraw(context)


# =====================================================
# REGISTER
# =====================================================

classes = (
    TesisPreferences,

    CaseItem,
    LandmarkItem,
    TesisState,

    TESIS_UL_Cases,

    TESIS_OT_EnsureUnits,
    TESIS_OT_ToggleCaseNeck,
    TESIS_OT_GenerateErrorCohort,
    TESIS_OT_ResetFilters,
    TESIS_OT_LoadExcel,
    TESIS_OT_ImportMeshes,
    TESIS_OT_LoadSavedLandmarks,
    TESIS_OT_AddLandmark,
    TESIS_OT_RefreshBVH,
    TESIS_OT_MarkNeckPoint,
    TESIS_OT_DiagnoseCoordinateFrame,
    TESIS_OT_ResetLandmark,
    TESIS_OT_SetLandmarkStatus,
    TESIS_OT_MarkLandmark,
    TESIS_OT_RecalculateActiveSOFT,
    TESIS_OT_RecalculateAllSOFT,
    TESIS_OT_CreateCCURing,
    TESIS_OT_ConfirmCCU,
    TESIS_OT_SaveCase,
    TESIS_OT_SaveCaseMetadata,
    TESIS_OT_ClearCurrentCase,
    TESIS_OT_CleanScene,
    TESIS_OT_NextCase,
    TESIS_OT_PrevCase,
    TESIS_OT_SaveAndNext,
    TESIS_OT_DiscardAndNext,
    TESIS_OT_NextLandmark,
    TESIS_OT_PrevLandmark,
    TESIS_OT_NextMissingLandmark,
    TESIS_OT_ToggleLandmarkNames,
    TESIS_OT_ToggleHardVisibility,
    TESIS_OT_ToggleSoftVisibility,
    TESIS_OT_ToggleLinesVisibility,
    TESIS_OT_SelectLandmark,
    TESIS_OT_DeleteSelected,
    TESIS_OT_RepairLandmarkLinks,
    TESIS_OT_SyncSelectedLandmark,
    TESIS_OT_MoveLandmarkSurface,
    TESIS_OT_SectionInfo,

    TESIS_PT_MainPanel,
    TESIS_PT_CasoPanel,
    TESIS_PT_ImportPanel,
    TESIS_PT_LandmarksPanel,
    TESIS_PT_CuelloPanel,
    TESIS_PT_GuardarPanel,
    TESIS_PT_HerramientasPanel,
)

_addon_keymaps = []

# Mapa obj.name → (lm_index, "SOFT"|"HARD") para lookup O(1) en handler
_OBJ_NAME_TO_LM: dict[str, tuple[int, str]] = {}

def _register_lm_objs(idx: int, soft_name: str, hard_name: str = ""):
    """Registra objetos y guarda metadatos persistentes para resolverlos tras recargas."""
    if soft_name:
        _OBJ_NAME_TO_LM[soft_name] = (idx, "SOFT")
        obj = bpy.data.objects.get(soft_name)
        if obj is not None:
            obj["tesis_lm_index"] = int(idx)
            obj["tesis_lm_role"] = "SOFT"
    if hard_name:
        _OBJ_NAME_TO_LM[hard_name] = (idx, "HARD")
        obj = bpy.data.objects.get(hard_name)
        if obj is not None:
            obj["tesis_lm_index"] = int(idx)
            obj["tesis_lm_role"] = "HARD"


def _unregister_lm_objs(idx: int):
    """Elimina todas las entradas del mapa que apunten al landmark idx."""
    to_del = [k for k, (i, _) in _OBJ_NAME_TO_LM.items() if i == idx]
    for k in to_del:
        del _OBJ_NAME_TO_LM[k]


def _resolve_lm_entry_for_object(state, obj):
    """Resuelve un objeto LM incluso si el mapa temporal se perdió al recargar."""
    entry = _OBJ_NAME_TO_LM.get(obj.name)
    if entry is not None:
        return entry

    try:
        idx = int(obj.get("tesis_lm_index", -1))
        role = _safe_str(obj.get("tesis_lm_role", "")).upper()
        if 0 <= idx < len(state.landmarks) and role in {"SOFT", "HARD"}:
            _OBJ_NAME_TO_LM[obj.name] = (idx, role)
            return idx, role
    except Exception:
        pass

    # Fallback definitivo: comparar con los nombres almacenados en TesisState.
    for idx, lm in enumerate(state.landmarks):
        if getattr(lm, "soft_obj_name", "") == obj.name:
            _register_lm_objs(idx, obj.name, getattr(lm, "hard_obj_name", ""))
            return idx, "SOFT"
        if getattr(lm, "hard_obj_name", "") == obj.name:
            _register_lm_objs(idx, getattr(lm, "soft_obj_name", ""), obj.name)
            return idx, "HARD"
    return None


def _set_object_world_location(obj, world_pos):
    """Asigna una posición WORLD correctamente aun si el objeto tiene parent."""
    world_pos = Vector(world_pos)
    if obj.parent is not None:
        obj.location = obj.parent.matrix_world.inverted_safe() @ world_pos
    else:
        obj.location = world_pos
    obj.update_tag(refresh={'OBJECT'})


_handler_surface_snap_guard = False
_undo_in_progress = False


def _reset_runtime_state():
    """Limpia estado temporal y referencias que pueden invalidarse con Undo/Redo."""
    global _marking_modal_active, _handler_surface_snap_guard, _undo_in_progress
    _marking_modal_active = False
    _handler_surface_snap_guard = False
    _undo_in_progress = False
    _OBJ_NAME_TO_LM.clear()
    invalidate_mallas_cache()
    invalidate_bvh_cache()
    try:
        _MALLAS_CACHE["key"] = None
        _MALLAS_CACHE["time"] = 0.0
        _MALLAS_CACHE["result"] = (None, None)
    except Exception:
        pass
    try:
        _clear_view3d_headers()
    except Exception:
        pass


def _sync_landmark_object_to_surface(scene, depsgraph, obj_name, *, show_header=True):
    """Compatibilidad: sincroniza el landmark seleccionado sin proyectarlo."""
    state = scene.tesis_state
    obj = bpy.data.objects.get(obj_name)
    if obj is None:
        return False
    entry = _resolve_lm_entry_for_object(state, obj)
    if entry is None:
        return False
    idx, _role = entry
    ok = _sync_landmark_data_from_objects(state, idx)
    if ok and show_header:
        screen = bpy.context.screen
        if screen:
            for area in screen.areas:
                if area.type == 'VIEW_3D':
                    area.header_text_set(f"{state.landmarks[idx].name} sincronizado")
                    area.tag_redraw()
    return ok


def _tesis_undo_pre(_dummy):
    """Evita usar referencias temporales mientras Blender restaura el undo."""
    global _undo_in_progress
    _undo_in_progress = True
    _OBJ_NAME_TO_LM.clear()
    invalidate_bvh_cache()


def _tesis_undo_post(_dummy):
    """Undo/Redo invalida objetos RNA; reconstruir caches por nombre al volver."""
    global _undo_in_progress
    _undo_in_progress = False
    _OBJ_NAME_TO_LM.clear()
    invalidate_bvh_cache()
    try:
        _MALLAS_CACHE["key"] = None
        _MALLAS_CACHE["time"] = 0.0
        _MALLAS_CACHE["result"] = (None, None)
    except Exception:
        pass


def register():
    _reset_runtime_state()
    _register_viewport_overlay()
    for c in classes:
        bpy.utils.register_class(c)
    bpy.types.Scene.tesis_state = PointerProperty(type=TesisState)

    # Keymap: Shift+SPACE en 3D View → iniciar/reanudar marcado secuencial
    # (evita conflicto con SPACE nativo de Blender: play/search)
    wm = bpy.context.window_manager
    kc = wm.keyconfigs.addon
    if kc:
        # Si quedó una instancia previa en memoria, remover keymaps rastreados antes de recrearlos.
        for km, kmi in _addon_keymaps:
            try:
                km.keymap_items.remove(kmi)
            except Exception:
                pass
        _addon_keymaps.clear()

        km = kc.keymaps.new(name='3D View', space_type='VIEW_3D')
        kmi = km.keymap_items.new('tesis.add_landmark', 'SPACE', 'PRESS', shift=True)
        kmi.properties.sequential = True
        _addon_keymaps.append((km, kmi))

        # G / X / DEL se interceptan en Object Mode, antes que los atajos
        # genéricos del 3D View. Para objetos no-LM se conserva Blender nativo.
        km_obj = kc.keymaps.new(name='Object Mode', space_type='EMPTY')

        kmi_move = km_obj.keymap_items.new('tesis.move_landmark_surface', 'G', 'PRESS')
        _addon_keymaps.append((km_obj, kmi_move))

        kmi_delete = km_obj.keymap_items.new('tesis.delete_selected', 'X', 'PRESS')
        _addon_keymaps.append((km_obj, kmi_delete))
        kmi_delete_alt = km_obj.keymap_items.new('tesis.delete_selected', 'DEL', 'PRESS')
        _addon_keymaps.append((km_obj, kmi_delete_alt))

        # Shift+RET → Guardar y siguiente (el más usado al finalizar cada caso)
        kmi2 = km.keymap_items.new('tesis.save_and_next', 'RET', 'PRESS', shift=True)
        _addon_keymaps.append((km, kmi2))

        # G usa raycast directo; las líneas se actualizan al crear o mover landmarks.

    # No se modifica geometría desde depsgraph_update_post: Blender advierte que
    # cambiar datos usados por el viewport desde ese handler puede provocar crash.
    if _tesis_undo_pre not in bpy.app.handlers.undo_pre:
        bpy.app.handlers.undo_pre.append(_tesis_undo_pre)
    if _tesis_undo_post not in bpy.app.handlers.undo_post:
        bpy.app.handlers.undo_post.append(_tesis_undo_post)
    if _tesis_undo_pre not in bpy.app.handlers.redo_pre:
        bpy.app.handlers.redo_pre.append(_tesis_undo_pre)
    if _tesis_undo_post not in bpy.app.handlers.redo_post:
        bpy.app.handlers.redo_post.append(_tesis_undo_post)

    log.info("Addon tesisjo registrado correctamente.")


def unregister():
    _reset_runtime_state()
    _unregister_viewport_overlay()
    global _DEFERRED_CLEANUP_TIMER, _DEFERRED_CLEANUP_PENDING
    try:
        if bpy.app.timers.is_registered(_run_deferred_cleanup):
            bpy.app.timers.unregister(_run_deferred_cleanup)
    except Exception:
        pass
    _DEFERRED_CLEANUP_TIMER = False
    _DEFERRED_CLEANUP_PENDING = False

    # Quitar handlers de Undo/Redo.
    for handler_list, fn in (
        (bpy.app.handlers.undo_pre, _tesis_undo_pre),
        (bpy.app.handlers.undo_post, _tesis_undo_post),
        (bpy.app.handlers.redo_pre, _tesis_undo_pre),
        (bpy.app.handlers.redo_post, _tesis_undo_post),
    ):
        if fn in handler_list:
            handler_list.remove(fn)

    # Quitar keymaps
    for km, kmi in _addon_keymaps:
        try:
            km.keymap_items.remove(kmi)
        except Exception:
            pass
    _addon_keymaps.clear()

    for c in reversed(classes):
        try:
            bpy.utils.unregister_class(c)
        except Exception:
            pass
    if hasattr(bpy.types.Scene, "tesis_state"):
        del bpy.types.Scene.tesis_state

    # Cerrar y remover FileHandler si existe para liberar el fichero en uso
    try:
        global _LOG_FH
        if _LOG_FH is not None:
            try:
                log.removeHandler(_LOG_FH)
            except Exception:
                pass
            try:
                _LOG_FH.close()
            except Exception:
                pass
            _LOG_FH = None
    except Exception:
        pass

    log.info("Addon tesisjo desregistrado.")

if __name__ == "__main__":
    print("Tesis Landmark System: activa este addon desde Edit > Preferences > Add-ons.")