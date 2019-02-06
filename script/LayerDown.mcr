current_layer = Application.ActiveDocument.ActiveLayer

if current_layer >= Application.ActiveDocument.ElectricalLayerCount then

    Application.ActiveDocument.ActiveLayer = 1

else

    Application.ActiveDocument.ActiveLayer = current_layer + 1

end if
