VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Documents_Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

private db		as ADODB.Connection


private sub Class_Initialize()

	set db = Currentproject.Connection
	
end sub


public function getDocument(docId as long) as Document


end function