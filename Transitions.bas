Attribute VB_Name = "Transitions"
'Module Functions...



  
            ' ******************************************************************************************************************************
            ' * procedure name: TransitionCLSIDToFriendlyName
            ' * procedure description: returns the localized friendly name of a transition given it's CLSID
            ' *
            ' ******************************************************************************************************************************
             Public Function TransitionCLSIDToFriendlyName(bstrTransitionCLSID As String, Optional bstrLanguage As String = "EN-US") As String
             Dim bstrReturn As String
             On Local Error GoTo ErrLine
             
             If UCase(bstrLanguage) = "EN-US" Then
                         Select Case bstrTransitionCLSID
                       
                            Case "{E31E87C4-86EA-4940-9B8A-5BD5D179A737}"
                                bstrReturn = "Basics"
                            Case "{C3BDF740-0B58-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Barn"
                            Case "{2E7700B7-27C4-437F-9FBF-1E8BE2817566}"
                                bstrReturn = "Bars"
                            Case "{00C429C0-0BA9-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Blinds"
                            Case "{107045D1-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Burn Film"
                            Case "{AA0D4D0C-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "CenterCurls"
                            Case "{2A54C908-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "ColorFade"
                            Case "{B3EE7802-8224-4787-A1EA-F0DE16DEABD3}"
                                bstrReturn = "Checkerboard"
                            Case "{9A43A844-0831-11D1-817F-0000F87557DB}"
                                bstrReturn = "Compositor"
                            Case "{AA0D4D0E-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "Curls"
                            Case "{AA0D4D12-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "Curtains"
                            Case "{F7F4A1B6-8E87-452f-A2D7-3077F508DBC0}"
                                bstrReturn = "Disolve"
                            Case "{16B280C5-EE70-11D1-9066-00C04FD9189D}"
                                bstrReturn = "Fade"
                            Case "{107045CC-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "FadeWhite"
                            Case "{2A54C90B-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "FlowMotion"
                            Case "{2A54C913-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "GlassBlock"
                            Case "{93073C40-0BA5-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Inset"
                            Case "{3F69F351-0379-11D2-A484-00C04F8EFB69}"
                                bstrReturn = "Iris"
                            Case "{2A54C904-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Jaws"
                            Case "{107045CA-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Lens"
                            Case "{107045C8-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "LightWipe"
                            Case "{AA0D4D0A-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "Liquid"
                            Case "{AA0D4D08-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "PageCurl"
                            Case "{AA0D4D10-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "PeelABCD"
                            Case "{4CCEA634-FBE0-11d1-906A-00C04FD9189D}"
                                bstrReturn = "Pixelate"
                            Case "{424B71AF-0695-11D2-A484-00C04F8EFB69}"
                                bstrReturn = "RadialWipe"
                            Case "{AA0D4D03-06A3-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "Ripple"
                            Case "{9C61F46E-0530-11D2-8F98-00C04FB92EB7}"
                                bstrReturn = "RollDown"
                            Case "{810E402F-056B-11D2-A484-00C04F8EFB69}"
                                bstrReturn = "Slide"
                            Case "{dE75D012-7A65-11D2-8CEA-00A0C9441E20}"
                                bstrReturn = "SMPTE Wipe"
                            Case "{ACA97E00-0C7D-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Spiral"
                            Case "{7658F2A2-0A83-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Stretch"
                            Case "{2A54C915-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Threshold"
                            Case "{63A4B1FC-259A-4A5B-8129-A83B8C9E6F4F}"
                                bstrReturn = "Strips"
                            Case "{107045CF-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Twister"
                            Case "{2A54C90D-07AA-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Vacuum"
                            Case "{107045C5-06E0-11D2-8D6D-00C04F8EF8E0}"
                                bstrReturn = "Water"
                            Case "{5AE1DAE0-1461-11d2-A484-00C04F8EFB69}"
                                bstrReturn = "Wheel"
                            Case "{AF279B30-86EB-11D1-81BF-0000F87557DB}"
                                bstrReturn = "Wipe"
                            Case "{0E6AE022-0C83-11D2-8CD4-00104BC75D9A}"
                                bstrReturn = "WormHole"
                            'Case "{E6E73D20-0C8A-11d2-A484-00C04F8EFB69}"
                             Case "{23E26328-3928-40F2-95E5-93CAD69016EB}"
                                bstrReturn = "Zigzag"
            
                                                        
                            Case Else: bstrReturn = vbNullString
                        End Select
            End If
            'return friendly name to the client
            
            TransitionCLSIDToFriendlyName = bstrReturn
            Exit Function
            
ErrLine:
            Err.Clear
            Exit Function
            End Function
             
             
             
            ' ******************************************************************************************************************************
            ' * procedure name: TransitionFriendlyNameToCLSID
            ' * procedure description: returns the CLSID of a transition given it's localized friendly name
            ' *
            ' ******************************************************************************************************************************
             Public Function TransitionFriendlyNameToCLSID(bstrFriendlyName As String, Optional bstrLanguage As String = "EN-US") As String
             Dim bstrReturn As String
             On Local Error GoTo ErrLine
             
             If UCase(bstrLanguage) = "EN-US" Then
                        
                        
                        Select Case bstrFriendlyName
                           
                            Case "Barn"
                                      bstrReturn = "{C3BDF740-0B58-11d2-A484-00C04F8EFB69}"
                            Case "Bars"
                                      bstrReturn = "{2E7700B7-27C4-437F-9FBF-1E8BE2817566}"
                            Case "Basics"
                                      bstrReturn = "{E31E87C4-86EA-4940-9B8A-5BD5D179A737}"
                            Case "Blinds"
                                      bstrReturn = "{00C429C0-0BA9-11d2-A484-00C04F8EFB69}"
                            Case "Burn Film"
                                      bstrReturn = "{107045D1-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "CenterCurls"
                                      bstrReturn = "{AA0D4D0C-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "Checkerboard"
                                      bstrReturn = "{B3EE7802-8224-4787-A1EA-F0DE16DEABD3}"
                            Case "ColorFade"
                                      bstrReturn = "{2A54C908-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Compositor"
                                      bstrReturn = "{9A43A844-0831-11D1-817F-0000F87557DB}"
                            Case "Curls"
                                      bstrReturn = "{AA0D4D0E-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "Curtains"
                                      bstrReturn = "{AA0D4D12-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "Disolve"
                                      bstrReturn = "{F7F4A1B6-8E87-452f-A2D7-3077F508DBC0}"
                            Case "Fade"
                                      bstrReturn = "{16B280C5-EE70-11D1-9066-00C04FD9189D}"
                            Case "FadeWhite"
                                      bstrReturn = "{107045CC-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "FlowMotion"
                                      bstrReturn = "{2A54C90B-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "GlassBlock"
                                      bstrReturn = "{2A54C913-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Grid"
                                      bstrReturn = "{2A54C911-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Inset"
                                      bstrReturn = "{93073C40-0BA5-11d2-A484-00C04F8EFB69}"
                            Case "Iris"
                                      bstrReturn = "{3F69F351-0379-11D2-A484-00C04F8EFB69}"
                            Case "Jaws"
                                      bstrReturn = "{2A54C904-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Lens"
                                      bstrReturn = "{107045CA-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "LightWipe"
                                      bstrReturn = "{107045C8-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "Liquid"
                                      bstrReturn = "{AA0D4D0A-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "PageCurl"
                                      bstrReturn = "{AA0D4D08-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "PeelABCD"
                                      bstrReturn = "{AA0D4D10-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "Pixelate"
                                      bstrReturn = "{4CCEA634-FBE0-11d1-906A-00C04FD9189D}"
                            Case "RadialWipe"
                                      bstrReturn = "{424B71AF-0695-11D2-A484-00C04F8EFB69}"
                            Case "Ripple"
                                      bstrReturn = "{AA0D4D03-06A3-11D2-8F98-00C04FB92EB7}"
                            Case "RollDown"
                                      bstrReturn = "{9C61F46E-0530-11D2-8F98-00C04FB92EB7}"
                            Case "Slide"
                                      bstrReturn = "{810E402F-056B-11D2-A484-00C04F8EFB69}"
                            Case "SMPTE Wipe"
                                      bstrReturn = "{dE75D012-7A65-11D2-8CEA-00A0C9441E20}"
                            Case "Spiral"
                                      bstrReturn = "{ACA97E00-0C7D-11d2-A484-00C04F8EFB69}"
                            Case "Stretch"
                                      bstrReturn = "{7658F2A2-0A83-11d2-A484-00C04F8EFB69}"
                            Case "Strips"
                                      bstrReturn = "{63A4B1FC-259A-4A5B-8129-A83B8C9E6F4F}"
                            Case "Threshold"
                                      bstrReturn = "{2A54C915-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Twister"
                                      bstrReturn = "{107045CF-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "Vacuum"
                                      bstrReturn = "{2A54C90D-07AA-11D2-8D6D-00C04F8EF8E0}"
                            Case "Water"
                                      bstrReturn = "{107045C5-06E0-11D2-8D6D-00C04F8EF8E0}"
                            Case "Wheel"
                                      bstrReturn = "{5AE1DAE0-1461-11d2-A484-00C04F8EFB69}"
                            Case "Wipe"
                                      bstrReturn = "{AF279B30-86EB-11D1-81BF-0000F87557DB}"
                            Case "WormHole"
                                      bstrReturn = "{0E6AE022-0C83-11D2-8CD4-00104BC75D9A}"
                            Case "Zigzag"
                                      bstrReturn = "{23E26328-3928-40F2-95E5-93CAD69016EB}"
                           
                            Case Else: bstrReturn = vbNullString
                        End Select
    '        bstrReturn = TransitionFriendlyNameToProgID(bstrFriendlyName)
            End If
            'return friendly name to the client
            
            TransitionFriendlyNameToCLSID = bstrReturn
            
            Exit Function
            
ErrLine:
            Err.Clear
            Exit Function
            End Function
            
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: EffectCLSIDToFriendlyName
            ' * procedure description: returns the localized friendly name of an effect given it's CLSID
            ' *
            ' ******************************************************************************************************************************
             Public Function EffectCLSIDToFriendlyName(bstrTransitionCLSID As String, Optional bstrLanguage As String = "EN-US") As String
             Dim bstrReturn As String
             On Local Error GoTo ErrLine
             
             If UCase(bstrLanguage) = "EN-US" Then
                         Select Case bstrTransitionCLSID
                            
                            Case "{BDD307C3-7BC0-4542-9F8F-A9611FE6C1BF}"
                                     bstrReturn = "Additive"
                            Case "{16B280C8-EE70-11D1-9066-00C04FD9189D}"
                                     bstrReturn = "BasicImage"
                            Case "{7312498D-E87A-11d1-81E0-0000F87557DB}"
                                     bstrReturn = "Blur"
                            Case "{5A20FD6F-F8FE-4a22-9EE7-307D72D09E6E}"
                                     bstrReturn = "Brightness"
                            Case "{421516C1-3CF8-11D2-952A-00C04FA34F05}"
                                     bstrReturn = "Chroma"
                            Case "{ADC6CB86-424C-11D2-952A-00C04FA34F05}"
                                     bstrReturn = "DropShadow"
                            Case "{F515306D-0156-11d2-81EA-0000F87557DB}"
                                     bstrReturn = "Emboss"
                            Case "{F515306E-0156-11d2-81EA-0000F87557DB}"
                                     bstrReturn = "Engrave"
                            Case "{16B280C5-EE70-11D1-9066-00C04FD9189D}"
                                     bstrReturn = "Fade"
                            Case "{9F8E6421-3D9B-11D2-952A-00C04FA34F05}"
                                     bstrReturn = "Glow"
                            Case "{3A04D93B-1EDD-4f3f-A375-A03EC19572C4}"
                                     bstrReturn = "MaskFilter"
                            Case "{4CCEA634-FBE0-11d1-906A-00C04FD9189D}"
                                     bstrReturn = "Pixelate"
                            Case "{E71B4063-3E59-11D2-952A-00C04FA34F05}"
                                     bstrReturn = "Shadow"
                            Case "{ADC6CB88-424C-11D2-952A-00C04FA34F05}"
                                     bstrReturn = "Wave"
                            Case Else: bstrReturn = vbNullString
                            End Select
            End If
            'return friendly name to the client
            
            EffectCLSIDToFriendlyName = bstrReturn
            
            Exit Function
            
ErrLine:
            Err.Clear
            Exit Function
            End Function
             
              
              
            ' ******************************************************************************************************************************
            ' * procedure name: EffectFriendlyNameToCLSID
            ' * procedure description: returns the CLSID of an effect given it's localized friendly name
            ' *
            ' ******************************************************************************************************************************
             Public Function EffectFriendlyNameToCLSID(bstrFriendlyName As String, Optional bstrLanguage As String = "EN-US") As String
             Dim bstrReturn As String
             On Local Error GoTo ErrLine
             
             If UCase(bstrLanguage) = "EN-US" Then
                        Select Case bstrFriendlyName
                            
                            Case "Additive"
                                     bstrReturn = "{BDD307C3-7BC0-4542-9F8F-A9611FE6C1BF}"
                            Case "BasicImage"
                                     bstrReturn = "{16B280C8-EE70-11D1-9066-00C04FD9189D}"
                            Case "Blur"
                                     bstrReturn = "{7312498D-E87A-11d1-81E0-0000F87557DB}"
                            Case "Brightness"
                                     bstrReturn = "{5A20FD6F-F8FE-4a22-9EE7-307D72D09E6E}"
                            Case "Chroma"
                                     bstrReturn = "{421516C1-3CF8-11D2-952A-00C04FA34F05}"
                            Case "DropShadow"
                                     bstrReturn = "{ADC6CB86-424C-11D2-952A-00C04FA34F05}"
                            Case "Emboss"
                                     bstrReturn = "{F515306D-0156-11d2-81EA-0000F87557DB}"
                            Case "Engrave"
                                     bstrReturn = "{F515306E-0156-11d2-81EA-0000F87557DB}"
                            Case "Fade"
                                     bstrReturn = "{16B280C5-EE70-11D1-9066-00C04FD9189D}"
                            Case "Glow"
                                      bstrReturn = "{9F8E6421-3D9B-11D2-952A-00C04FA34F05}"
                            Case "MaskFilter"
                                      bstrReturn = "{3A04D93B-1EDD-4f3f-A375-A03EC19572C4}"
                            Case "Pixelate"
                                     bstrReturn = "{4CCEA634-FBE0-11d1-906A-00C04FD9189D}"
                            Case "Shadow"
                                      bstrReturn = "{E71B4063-3E59-11D2-952A-00C04FA34F05}"
                            Case "Wave"
                                     bstrReturn = "{ADC6CB88-424C-11D2-952A-00C04FA34F05}"
                          
                            Case Else: bstrReturn = vbNullString
                        End Select
            End If
            'return friendly name to the client
            EffectFriendlyNameToCLSID = bstrReturn
            Exit Function
            
ErrLine:
            Err.Clear
            Exit Function
            End Function

 Public Function TransitionFriendlyNameToProgID(bstrTransitionFriendlyName As String) As String
            On Local Error GoTo ErrLine
            
            Select Case LCase(Trim(bstrTransitionFriendlyName))
                Case "default"
                         TransitionFriendlyNameToProgID = "DxtJpegDll.DxtJpeg"
                Case "slide"
                         TransitionFriendlyNameToProgID = "DXImageTransform.Microsoft.CrSlide"
                Case "fade"
                         TransitionFriendlyNameToProgID = "DXImageTransform.Microsoft.Fade"
                Case "ripple"
                         TransitionFriendlyNameToProgID = "DXImageTransform.MetaCreations.Water"
                Case "circle"
                         TransitionFriendlyNameToProgID = "DXImageTransform.MetaCreations.Grid"
                Case "burn film"
                         TransitionFriendlyNameToProgID = "DXImageTransform.MetaCreations.BurnFilm"
                Case "barn doors"
                         TransitionFriendlyNameToProgID = "DXImageTransform.Microsoft.CrBarn"
            End Select
            Exit Function
            
ErrLine:
            Err.Clear
            Exit Function
            End Function
