#FusionAPI_python ExportWire Ver0.0.2
#Author-kantoku
#Description-Export sketch lines/points.

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
        crvs,_ = tmpMgr.createWireFromCurves(geos, True)
        
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

#言語別メッセージ　日本語以外は英語
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

#ファイルエクスポート
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
    
#ﾌｧｲﾙパス
def Get_Filepath(ui):
    dlg = ui.createFileDialog()
    dlg.title = '3DCurvesExport'
    dlg.isMultiSelectEnabled = False
    dlg.filter = 'IGES(*.igs);;STEP(*.stp);;SAT(*.sat)'
    if dlg.showSave() != adsk.core.DialogResults.DialogOK :
        return
    return dlg.filename

#新しいDocs
def NewDoc(app):
    desDoc = adsk.core.DocumentTypes.FusionDesignDocumentType
    return app.documents.add(desDoc)

#World座標でのジオメトリ取得
def GetSketchCurvesGeos(skt):
    if len(skt.sketchCurves) < 1:
        return None
    
    #extension
    adsk.fusion.SketchCurve.toGeoTF = SketchCurveToGeoTransform
    adsk.fusion.Component.toOcc = ComponentToOccurrenc
    
    mat = skt.transform.copy()
    occ = skt.parentComponent.toOcc()
    
    if not occ is None:
        mat.transformBy(occ.transform)
        
    geos = [crv.toGeoTF(mat) for crv in skt.sketchCurves if not crv.isConstruction]
    
    return geos

#adsk.fusion.SketchCurve
def SketchCurveToGeoTransform(self,mat3d):
    geo = self.geometry.copy()
    geo.transformBy(mat3d)
    
    return geo

#adsk.fusion.Component 拡張メソッド
#コンポーネントからオカレンスの取得　ルートはNone
def ComponentToOccurrenc(self):
    root = self.parentDesign.rootComponent
    if self == root:
        return None
        
    occs = [occ
            for occ in root.allOccurrencesByComponent(self)
            if occ.component == self]
    return occs[0]