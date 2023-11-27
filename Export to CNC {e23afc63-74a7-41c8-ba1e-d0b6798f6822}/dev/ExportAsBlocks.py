#Convert all layers that are on to blocks and export as DWG\STP for solid modelers
import rhinoscriptsyntax as rs

def FusionExport():
    blocks = list()
    blockNames = list()
    point = [0,0,0]
    layers = rs.LayerIds()
    
    if layers:
        rs.UnselectAllObjects()
        rs.EnableRedraw(False)
        lower = 1
        upper = len(layers)
        rs.StatusBarProgressMeterShow("Making blocks", lower, upper)
        for layer in layers:
            if rs.IsLayer(layer):
                parts = rs.ObjectsByLayer(layer, False)
                if (parts) and (rs.IsLayerVisible(layer)):
                    #print rs.LayerName(layer, False)
                    rs.CurrentLayer(layer)
                    layer = rs.LayerName(layer).replace('::','_')
                    block = rs.AddBlock(parts, point, layer, False)
                    if (block):
                        blocks.append(rs.InsertBlock(block, point))
                        blockNames.append(block)
            rs.StatusBarProgressMeterUpdate(lower)
            lower += 1
            rs.Sleep(20)

        rs.SelectObjects(blocks)
        filename = rs.SaveFileName ("Save",
        "STP Files (*.stp)|*.stp|DWG Files (*.dwg)|*.dwg|DXF Files (*.dxf)|*.dxf|All Files (*.*)|*.*||")
        if filename:
            rs.Command('_-Export "' + filename + '" _Enter', False)
        else:
            print('Export cancelled no file name selected')
        rs.StatusBarProgressMeterHide()
        lower = 1
        upper = len(blockNames)
        rs.StatusBarProgressMeterShow("Deleting blocks", lower, upper)
        for name in blockNames:
            rs.DeleteBlock(name)
            rs.StatusBarProgressMeterUpdate(lower)
            lower += 1
            rs.Sleep(25)
        rs.StatusBarProgressMeterHide()
        rs.EnableRedraw(True)

##########################################################################
# Here we check to see if this file is being executed as the "main" python
# script instead of being used as a module by some other python script
# This allows us to use the module which ever way we want.
if( __name__ == '__main__' ):
    FusionExport()
