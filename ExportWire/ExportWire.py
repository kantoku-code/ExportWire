#FusionAPI_python ExportWire Ver0.0.3
#Author-kantoku
#Description-表示されている全てのスケッチの線をエクスポート
#コンストラクションはエクスポートしません

import adsk.core, adsk.fusion, traceback
from itertools import chain
import os.path

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument
        des = adsk.fusion.Design.cast(app.activeProduct)
        msg_dic = GetMsg()
        
        #表示されているスケッチ
        skts = [skt
                for comp in des.allComponents if comp.isSketchFolderLightBulbOn
                for skt in comp.sketches if skt.isVisible]
        ui.activeSelections.clear()
        
        #正しい位置でｼﾞｵﾒﾄﾘ取得       
        geos = list(chain.from_iterable(GetSketchCurvesGeos(skt) for skt in skts))
        if len(geos) < 1:
            ui.messageBox(msg_dic['not_found'])
            return
        
        #エクスポートﾌｧｲﾙﾊﾟｽ
        path = Get_Filepath(ui)
        if path is None:
            return
        
        #新規デザイン
        expDoc = NewDoc(app)
        expDes = adsk.fusion.Design.cast(app.activeProduct)
        doc.activate()
        
        #ダイレクト
        expDes.designType = adsk.fusion.DesignTypes.DirectDesignType

        #tempBRep
        tmpMgr = adsk.fusion.TemporaryBRepManager.get()
        crvs, _ = tmpMgr.createWireFromCurves(geos, True)
        
        #実体化
        expRoot = expDes.rootComponent
        bodies = expRoot.bRepBodies
        bodies.add(crvs)
        
        #保存
        res = ExportFile(path,expDes.exportManager)
        
        #一時Docを閉じる
        expDoc.close(False)
        
        #おしまい
        if res:
            msg = msg_dic['done']
        else:
            msg = msg_dic['failed']
        
        ui.messageBox(msg)
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

#lang message
def GetMsg():
    langs = adsk.core.UserLanguages
    
    app = adsk.core.Application.get()
    lang = app.preferences.generalPreferences.userLanguage
    
    keys = ['not_found','done','failed']
    if lang == langs.JapaneseLanguage:
        values = ['エクスポートする要素がありません!','完了','失敗しました!']
    else:
        values = ['Export item not found!', 'Done', 'Failed!']
    
    return dict(zip(keys, values))

#file ExportOptions
def ExportFile(path,expMgr):
    _, ext = os.path.splitext(path)
    
    if 'igs' in ext:
        expOpt = expMgr.createIGESExportOptions(path)
    elif 'stp' in ext:
        expOpt = expMgr.createSTEPExportOptions(path)
    elif 'sat' in ext:
        expOpt = expMgr.createSATExportOptions(path)
    else:
        return False
        
    expMgr.execute(expOpt)
    return True
    
#filepath Dialog
def Get_Filepath(ui):
    dlg = ui.createFileDialog()
    dlg.title = '3DCurvesExport'
    dlg.isMultiSelectEnabled = False
    dlg.filter = 'IGES(*.igs);;STEP(*.stp);;SAT(*.sat)'
    if dlg.showSave() != adsk.core.DialogResults.DialogOK :
        return
    return dlg.filename

#new Doc    
def NewDoc(app):
    desDoc = adsk.core.DocumentTypes.FusionDesignDocumentType
    return app.documents.add(desDoc)

#Root Geometry
def GetSketchCurvesGeos(skt):
    if len(skt.sketchCurves) < 1:
        return None
    
    #extension
    adsk.fusion.SketchCurve.toGeoTF = SketchCurveToGeoTransform
    adsk.fusion.Component.rootMatrix = GetRootMatrix
    
    #Matrix
    mat = skt.parentComponent.rootMatrix()

    #Geometry Transform
    geos = [crv.toGeoTF(mat) for crv in skt.sketchCurves if not crv.isConstruction]
    
    return geos

#adsk.fusion.SketchCurve  extension_method
def SketchCurveToGeoTransform(self, mat3d):
    geo = self.worldGeometry.copy()
    geo.transformBy(mat3d)
    
    return geo

#adsk.fusion.Componen extension_method
def GetRootMatrix(self):
    comp = adsk.fusion.Component.cast(self)
    des = adsk.fusion.Design.cast(comp.parentDesign)
    root = des.rootComponent

    mat = adsk.core.Matrix3D.create()
  
    if comp == root:
        return mat

    occs = root.allOccurrencesByComponent(comp)
    if len(occs) < 1:
        return mat
    
    occ = occs[0]
    occ_names = occ.fullPathName.split('+')
    occs = [root.allOccurrences.itemByName(name) 
                for name in occ_names]
    mat3ds = [occ.transform for occ in occs]
    mat3ds.reverse()
    for mat3d in mat3ds:
        mat.transformBy(mat3d)               

    return mat