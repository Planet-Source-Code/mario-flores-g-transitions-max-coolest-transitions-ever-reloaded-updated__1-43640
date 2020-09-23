Attribute VB_Name = "Variables"
Public Enum StyleSlide
    XEffect = 0
    XTransition = 1
End Enum
Public i
Public SourcePath   As String
Public Style As StyleSlide
Public Pfiles As Collection
Public WindowFlag As Boolean
Public SLIDETIME As Integer
Global gbl_objTimeline As AMTimeline
Public bstrCurrentFile As String
Public PictureDelayStart As Double
Public PictureDelayEnd As Double
Public SlidePicture As AMTimeline
Public PictureSource As AMTimelineSrc
Public NewPicture As AMTimelineTrack
Public PictureTransition As AMTimelineTrans
Public Transition As IAMTimelineTransable
Public Slide As AMTimelineGroup
Public CurrentPicture As AMTimelineObj
Public Node As AMTimelineObj
Public NodeGroup As AMTimelineObj
Public PictureGroup As AMTimelineComp
Public Effect As AMTimelineEffect
Public NodeEffect As IAMTimelineEffectable
Public nResultant As Long
Public objPosition As IMediaPosition
Public objMediaEvent As IMediaEvent
Public objFilterGraph As IGraphBuilder
Public objVideoWindow As IVideoWindow
Public objRenderEngine As RenderEngine
Public objFilterGraphManager As New FilgraphManager
Public FullSize As Boolean
Public Paused As Boolean
Public Running As Boolean
