Attribute VB_Name = "SlideShow"

Option Explicit
Option Base 0
Option Compare Text

'user for async rendering procedures
Private m_objMediaEvent As IMediaEvent
Private m_objFilterGraph As IGraphBuilder
Private m_objRenderEngine As RenderEngine
Private m_objFilterGraphManager As New FilgraphManager
Dim Z As Variant
Dim nCount As Long

            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'THIS SUB CREATES THE SLIDES IN THE SLIDESHOW
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


Public Function CreateSLIDESHOW(DefaultTransitionCLSID As String, files As Collection) As AMTimeline
            On Local Error GoTo ErrLine
            
            
'****************************************************************************************
        'CREATE SLIDE PRESENTATION(EMPTY NODE)
            
            Set SlidePicture = New AMTimeline
            Call SlidePicture.EnableTransitions(1)
            Call SlidePicture.EnableEffects(1)
            SlidePicture.CreateEmptyNode Node, TIMELINE_MAJOR_TYPE_GROUP
            Set Slide = Node
            Call Slide.SetGroupName("SLIDEMarioFlores")
            Call Slide.SetMediaTypeForVB(0)
            Set Node = Slide
            SlidePicture.AddGroup Node
            Set NodeGroup = Slide
'****************************************************************************************


nCount = -1
For Each Z In files
                 

'****************************************************************************************
        'COUNT PICTURES IN SLIDE PRESENTATION
         nCount = nCount + 1
         bstrCurrentFile = Z(0)
'****************************************************************************************

'****************************************************************************************
        'INSERT NEW PICTURE INTO SLIDE PRESENTATION
                   
        SlidePicture.CreateEmptyNode Node, TIMELINE_MAJOR_TYPE_TRACK
        Set NewPicture = Node
        Set CurrentPicture = NewPicture
        Set PictureGroup = NodeGroup
        Call PictureGroup.VTrackInsBefore(Node, -1)
        
'****************************************************************************************

'****************************************************************************************
        'CREATE SOURCE OF NEW PICTURE IN SLIDE PRESENTATION (PLACE OF PICURE IN SLIDE)
        
        SlidePicture.CreateEmptyNode Node, TIMELINE_MAJOR_TYPE_SOURCE
        Set PictureSource = Node
         
'****************************************************************************************
            
                         
'****************************************************************************************
        'CALCULATE DELAY TIME OF PICTURES DURING SLIDE PRESENTATION
                           
        If PictureDelayEnd = 0 Then
           PictureDelayStart = SLIDETIME * (nCount)
           PictureDelayEnd = (SLIDETIME * (nCount + 1)) + 1
        
        Else
           PictureDelayStart = (SLIDETIME * (nCount)) - 1
           PictureDelayEnd = (SLIDETIME * (nCount + 1)) + 1
           
        End If

'****************************************************************************************
                           
'****************************************************************************************
       'SET SOURCE OF NEW PICTURE OF SLIDE PRESENTATION (PLACE OF PICURE IN SLIDE)
                 
        Call PictureSource.SetMediaName(bstrCurrentFile)
        Set Node = PictureSource
        Call Node.SetStartStop2(PictureDelayStart, PictureDelayEnd)
        NewPicture.SrcAdd Node
               
        
            
             
'****************************************************************************************
        'SET TRANSITION & EFFECT IN EACH PICTURE OF SLIDE PRESENTATION
               
        If Style = XEffect Then
           SlidePicture.CreateEmptyNode Node, TIMELINE_MAJOR_TYPE_EFFECT
           Set Effect = Node
        End If
        
        If Style = XTransition Then
            SlidePicture.CreateEmptyNode Node, TIMELINE_MAJOR_TYPE_TRANSITION
            
            Set PictureTransition = Node
        End If
        
        PictureDelayStart = (SLIDETIME * nCount) - 1
        PictureDelayEnd = SLIDETIME * nCount + 1
        
        If PictureDelayStart < 0 Then PictureDelayStart = 0
        
        If Style = XEffect Then
             Set Node = Effect
             Call Node.SetSubObjectGUIDB(DefaultTransitionCLSID)
             Call Node.SetStartStop2(PictureDelayStart, PictureDelayEnd)
             Set NodeEffect = CurrentPicture
             Call NodeEffect.EffectInsBefore(Node, -1)
        End If
        If Style = XTransition Then
            Set Node = PictureTransition
            Call Node.SetSubObjectGUIDB(DefaultTransitionCLSID)
            Call Node.SetStartStop2(PictureDelayStart, PictureDelayEnd)
            Set Transition = CurrentPicture
            Call Transition.TransAdd(Node)
        End If
       
Next Z
        'Destroy Objects...
        'No need to use them anymore because SLIDE PRESENTATION is already created

         Set CreateSLIDESHOW = SlidePicture
         Set SlidePicture = Nothing
         Set PictureSource = Nothing
         Set NewPicture = Nothing
         Set PictureTransition = Nothing
         Set Transition = Nothing
         Set Effect = Nothing
         Set PictureGroup = Nothing
         Set Slide = Nothing
         Set CurrentPicture = Nothing
         Set Node = Nothing
         Set NodeGroup = Nothing
         Set NodeEffect = Nothing
         PictureDelayEnd = 0
         Exit Function
'****************************************************************************************
ErrLine:
         Err.Clear
         Exit Function
         End Function







'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'THIS SUB DISPLAYS THE SLIDES THAT ARE ALLREADY CREATED
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
       
Public Sub SHOWALL(SlidePicture As AMTimeline)

On Local Error GoTo ErrLine
Dim NumFotos
    NumFotos = 0

If Not SlidePicture Is Nothing Then
               
               frmMain.Espere.Visible = True
               frmMain.Refresh
' ******************************************************************************************************************************
               'CONNECT EVERYTHING
            
               Set objRenderEngine = New RenderEngine
               
               Call objRenderEngine.SetTimelineObject(SlidePicture)
               objRenderEngine.ConnectFrontEnd
               objRenderEngine.RenderOutputPins
' ******************************************************************************************************************************
               
' ******************************************************************************************************************************
               'LET DIRECTDRAW RENDER VIDEO
               Call objRenderEngine.GetFilterGraph(objFilterGraph)
               Set objFilterGraphManager = New FilgraphManager
               Set objFilterGraphManager = objFilterGraph
               Set objPosition = objFilterGraphManager
' ******************************************************************************************************************************
              
              
               
' ******************************************************************************************************************************
               'CREATE DISPLAY WINDOW
               Set objMediaEvent = objFilterGraphManager
               Set objVideoWindow = objMediaEvent
               
               
               'Size on Screen (RECTANGLE)(STAND ALONE)
               'objVideoWindow.Width = (Screen.Width / Screen.TwipsPerPixelX) / 2
               'objVideoWindow.Height = (Screen.Width / Screen.TwipsPerPixelY) / 3
               'Center on Screen(Stand Alone)
               'objVideoWindow.Left = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - (objVideoWindow.Width / 2)
               'objVideoWindow.Left =((Screen.Height / Screen.TwipsPerPixelY) / 2) - (objVideoWindow.Height / 2)
               
               'Size on Screen (RECTANGLE)
               objVideoWindow.Width = (frmMain.Frame.Width / Screen.TwipsPerPixelX)
               objVideoWindow.Height = (frmMain.Frame.Height / Screen.TwipsPerPixelY)
               
                
               'Center on (FORM)
            
               objVideoWindow.Left = (frmMain.Frame.Left / Screen.TwipsPerPixelX) + (frmMain.Left / Screen.TwipsPerPixelX)
               objVideoWindow.Top = (frmMain.Frame.Top / Screen.TwipsPerPixelY) + (frmMain.Top / Screen.TwipsPerPixelY)
               
               objVideoWindow.Caption = "MArio Flores G..."
               objVideoWindow.WindowStyle = &H40000000 Or &H10000000 'CHILD + VISIBLE
               objVideoWindow.WindowState = 1 'FULL SCREEN=3
               objVideoWindow.WindowStyleEx = 1 'ALWAYS ON TOP= -1
               objFilterGraphManager.Run 'Here We GO!!!
               frmMain.Espere.Visible = False
' ******************************************************************************************************************************

            

' ******************************************************************************************************************************
               'DO LOOP UNTIL PRESENTATION FINISH
               Do: DoEvents
                     
               If Paused = True Then
               Call objFilterGraphManager.Pause
               End If
               If Paused = False Then
               Call objFilterGraphManager.Run
               
               
               
               If Not objMediaEvent Is Nothing Then Call objMediaEvent.WaitForCompletion(10, nResultant)
                   Running = True
                   NumFotos = Round((((SLIDETIME * (nCount + 1)) - (objPosition.CurrentPosition))) / SLIDETIME) 'Dont Exactly know what i did!..Still Works
                   NumFotos = Abs(NumFotos - (nCount + 1))
                   If NumFotos <= nCount + 1 Then frmMain.lblTime = NumFotos & " / " & nCount + 1
                   If FullSize = True Then objVideoWindow.WindowState = 3

                     If nResultant = 1 Then
                          If Not objVideoWindow Is Nothing Then
                                 objVideoWindow.Left = Screen.Width * 8
                                 objVideoWindow.Top = Screen.Height * 8
                                 objVideoWindow.Visible = False
                                 frmMain.lblNumber.Caption = "Slide Show Completed " & nCount + 1 & " Photo(s) Displayed"
                                 ENDSHOW
                          End If
                          If Not objFilterGraphManager Is Nothing Then _
                                 Call objFilterGraphManager.Stop
                          Exit Do
                     ElseIf objVideoWindow.Visible = False Then
                          If Not objFilterGraphManager Is Nothing Then _
                                 Call objFilterGraphManager.Stop
                          Exit Do
                     ElseIf SlidePicture Is Nothing Then
                          Exit Do
                     ElseIf objFilterGraphManager Is Nothing Then
                          Exit Do
                     End If
                        

               
               End If
               Loop
               
            Else: nResultant = 1

End If
' ******************************************************************************************************************************
          'Clean Up
          Set objPosition = Nothing
          Set objFilterGraphManager = Nothing
          Set objMediaEvent = Nothing
          Set objVideoWindow = Nothing
          Set objRenderEngine = Nothing
          Set SlidePicture = Nothing
          Set objFilterGraph = Nothing
          Running = False
          frmMain.lblTime = vbNullString
          Unload ToolBox
' ******************************************************************************************************************************
          
            Exit Sub
' ******************************************************************************************************************************
ErrLine:
            Err.Clear
            Resume Next
            Exit Sub
            End Sub








'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'END SUB...CLOSE EVERYTHING
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
Public Sub ENDSHOW()

If WindowFlag = True Then Exit Sub


Call SHOWALL(Nothing)



End Sub
            
            
            
    
