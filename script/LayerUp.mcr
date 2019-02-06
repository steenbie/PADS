current_layer = Application.ActiveDocument.ActiveLayer

if current_layer <= 1 then

    Application.ActiveDocument.ActiveLayer = Application.ActiveDocument.ElectricalLayerCount

else

    Application.ActiveDocument.ActiveLayer = current_layer - 1

end if
