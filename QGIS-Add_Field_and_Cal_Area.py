#### เพิ่มพื้นที่คำนวณลงในทุกเลเยอร์ที่เป็นโพลิกอนในโปรเจกต์ QGIS

from qgis.core import (
    QgsProject,
    QgsField,
    QgsDistanceArea,
    QgsMapLayerType,
    QgsWkbTypes
)
from qgis.PyQt.QtCore import QVariant

field_name = 'Area_Calc'
project = QgsProject.instance()

# คำนวณพื้นที่โดยใช้ QgsDistanceArea ซึ่งจะคำนวณพื้นที่ตามระบบพิกัดของแต่ละเลเยอร์
# การใช้ QgsDistanceArea จะช่วยให้ได้ค่าพื้นที่ที่ถูกต้องแม้ในกรณีที่เลเยอร์มีระบบพิกัดที่แตกต่างกัน
# ปกติจะทำงานใน EPSG 24047 กับ 24048 สำหรับประเทศไทย
d = QgsDistanceArea()
d.setEllipsoid(project.ellipsoid())

for layer in project.mapLayers().values():

    # จะดูเฉพาะเลเยอร์ที่เป็นโพลิกอนเท่านั้น
    if layer.type() != QgsMapLayerType.VectorLayer:
        continue
    if layer.geometryType() != QgsWkbTypes.PolygonGeometry:
        continue

    print(f"กำลังประมวลผลเลเยอร์: {layer.name()}")

    # สำคัญมาก บรรทัดนี้จะทำให้การคำนวณพื้นที่ตามระบบพิกัดของเลเยอร์นั้นๆ
    d.setSourceCrs(layer.crs(), project.transformContext())

    layer.startEditing()

    # เพิ่มฟิลด์ใหม่ถ้ายังไม่มี
    if layer.fields().indexFromName(field_name) == -1:
        layer.dataProvider().addAttributes([
            QgsField(field_name, QVariant.Double, 'double', 20, 3)
        ])
        layer.updateFields()

    field_index = layer.fields().indexFromName(field_name)

    for feature in layer.getFeatures():
        geom = feature.geometry()

        # ใส่การตรวจสอบเพิ่มเติมเพื่อข้ามฟีเจอร์ที่รูปร่างมีปัญหา เช่น รูปร่างที่ไม่สมบูรณ์หรือมีข้อผิดพลาด
        if geom is None or geom.isEmpty():
            continue

        area_val = d.measureArea(geom)

        # ตรวจสอบค่า area_val ว่ามีค่าเป็น None หรือ NaN หรือไม่ เพื่อป้องกันการบันทึกค่าที่ไม่ถูกต้องลงในฟิลด์
        if area_val is None or area_val != area_val:
            continue

        layer.changeAttributeValue(
            feature.id(),
            field_index,
            round(area_val, 3)
        )

    layer.commitChanges()
    print(f"เพิ่ม Field {field_name} ในเลเยอร์ {layer.name()} สำเร็จ!")

print("--- ประมวลผลเสร็จสิ้นทุกเลเยอร์ ---")