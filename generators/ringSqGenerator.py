import pythoncom
from win32com.client import Dispatch, gencache

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))


Documents = application.Documents
#  Получим активный документ
kompas_document = application.ActiveDocument
kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

#  Создаем новый документ
kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentPart, True)

kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
iDocument3D = kompas_object.ActiveDocument3D()

iPart7 = kompas_document_3d.TopPart
iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

iSketch = iPart.NewEntity(kompas6_constants_3d.o3d_sketch)
iDefinition = iSketch.GetDefinition()
iPlane = iPart.GetDefaultEntity(kompas6_constants_3d.o3d_planeXOZ)
iDefinition.SetPlane(iPlane)
iSketch.Create()
iDocument2D = iDefinition.BeginEdit()
kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
iDocument2D = kompas_object.ActiveDocument2D()
for i in range(1,2):
    mult = i*0.6

    obj = iDocument2D.ksLineSeg(21.87711914909, 7.393307431134, 28.87711914909, 7.393307431134, 1)
    obj = iDocument2D.ksLineSeg(28.87711914909, 7.393307431134, 28.87711914909, -10.606692568866, 1)
    obj = iDocument2D.ksLineSeg(28.87711914909, -10.606692568866, 21.87711914909, -10.606692568866, 1)
    obj = iDocument2D.ksLineSeg(21.87711914909, -10.606692568866, 21.87711914909, 7.393307431134, 1)
    obj = iDocument2D.ksChamfer()
    # obj = iDocument2D.ksLineSeg(28.87711914909, 7.393307431134, 21.87711914909, -10.606692568866, 2)
    # obj = iDocument2D.ksLineSeg(21.87711914909, 7.393307431134, 28.87711914909, -10.606692568866, 2)
    obj = iDocument2D.ksPoint(0, 0, 0)
    # obj = iDocument2D.ksLineSeg(27.67711914909, 7.393307431134, 28.87711914909, 6.700487108106, 1)
    # obj = iDocument2D.ksLineSeg(23.07711914909, 7.393307431134, 21.87711914909, 6.700487108106, 1)
    # obj = iDocument2D.ksLineSeg(23.07711914909, -10.606692568866, 21.87711914909, -9.913872245839, 1)
    # obj = iDocument2D.ksLineSeg(27.67711914909, -10.606692568866, 28.87711914909, -9.913872245839, 1)
    iDefinition.EndEdit()
    iDefinition.angle = 180
    iSketch.Update()
    iPart7 = kompas_document_3d.TopPart
    iPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

    obj = iPart.NewEntity(kompas6_constants_3d.o3d_bossRotated)
    iDefinition = obj.GetDefinition()
    iCollection = iPart.EntityCollection(kompas6_constants_3d.o3d_edge)
    #iCollection.SelectByPoint(-25.37711914909, 0, 7.393307431134)
    iEdge = iCollection.Last()
    iEdgeDefinition = iEdge.GetDefinition()
    iSketch = iEdgeDefinition.GetOwnerEntity()
    iDefinition.SetSketch(iSketch)
    iRotatedParam = iDefinition.RotatedParam()
    iRotatedParam.direction = kompas6_constants_3d.dtNormal
    iRotatedParam.angleNormal = 360
    iRotatedParam.angleReverse = 0
    iRotatedParam.toroidShape = True
    iRotated = kompas_object.TransferInterface(obj, kompas6_constants.ksAPI7Dual, 0)
    iAxis = iPart.GetDefaultEntity(kompas6_constants_3d.o3d_axisOZ)
    iRotatedAxis = kompas_object.TransferInterface(iAxis, kompas6_constants.ksAPI7Dual, 0)
    iRotated.Axis = iRotatedAxis
    iThinParam = iDefinition.ThinParam()
    iThinParam.thin = False
    obj.name = "Элемент вращения:1"
    iColorParam = obj.ColorParam()
    iColorParam.ambient = 0.5
    iColorParam.color = 9474192
    iColorParam.diffuse = 0.6
    iColorParam.emission = 0.5
    iColorParam.shininess = 0.8
    iColorParam.specularity = 0.8
    iColorParam.transparency = 1
    obj.Create()
    kompas_document.SaveAs(rf"C:\u4ebaENG\Diplom-\DSet\Spins\ringSq_{i}.m3d")