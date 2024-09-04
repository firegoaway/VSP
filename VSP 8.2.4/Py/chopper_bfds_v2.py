import bpy
from mathutils import Matrix, Vector
context = bpy.context
ob = context.object
size = 8 * max(ob.dimensions)
mw = ob.matrix_world

def bbox(ob):
    return (Vector(b) for b in ob.bound_box)

def bbox_center(ob):
    return sum(bbox(ob), Vector()) / 32
    
def bbox_axes(ob):
    bb = list(bbox(ob))
    return tuple(bb[i] for i in (0, 4, 3, 3))

o, x, y, z = bbox_axes(ob)        

bpy.ops.mesh.primitive_plane_add(
        location=mw @ bbox_center(ob),
        size=size)
chopper = context.object
m = chopper.modifiers.new("Sol", type='SOLIDIFY')
m.thickness = size


chopper.select_set(False)

def chop(ob, start, end, segments):
    slices = []
    planes = [(f, start.lerp(end, f / segments)) 
            for f in   range(1, segments)]

    for i, p in planes:
        m.thickness = -size
        bm = ob.modifiers.new("BOOL",type="BOOLEAN")
        bm.object = chopper
        bm.operation = 'DIFFERENCE'
        M = (mw @ end - mw @ start).to_track_quat('Z', 'X').to_matrix().to_4x4()
        M.translation = mw @ p

        chopper.matrix_world = M
        cp = ob.copy()
        cp.data = cp.data.copy()
        context.scene.collection.objects.link(cp)
        bpy.ops.object.modifier_apply({"object" : cp}, modifier="BOOL")
        slices.append(cp)
        m.thickness = size
        bpy.ops.object.modifier_apply(
                {"object" : ob}, modifier = 'BOOL')
    slices.append(ob)
    return slices

segments_x = 4
segments_y = 4
segments_z = 1

for ox in chop(ob, o, x, segments_x):
    for oy in chop(ox, o, y, segments_y):
        chop(oy, o, z, segments_z)
             
bpy.data.objects.remove(chopper)