Attribute VB_Name = "modCommonDialog"
Option Explicit

Public Const PictureFilter = "Picture Files|*.bmp;*.ico;*.gif;*.jpg|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico|GIF Files (*.gif)|*.gif|JPEG Files (*.jpg)|*.jpg|All Files (*.*)|*.*"

Public Const MapFilter = "Evil Game Map (*.egm)|*.egm|All Files (*.*)|*.*"
Public Const TilesetsFilter = "Evil Game Tilesets (*.egt)|*.egt|All Files (*.*)|*.*"
Public Const AnimationsFilter = "Evil Game Animations (*.ega)|*.ega|All Files (*.*)|*.*"
Public Const ResourceFilter = "Resource Files|*.egt;*.ega|Evil Game Tilesets (*.egt)|*.egt|Evil Game Animations (*.ega)|*.ega|All Files (*.*)|*.*"

Public Const OpenFlags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
Public Const MultiOpenFlags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
Public Const SaveFlags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
