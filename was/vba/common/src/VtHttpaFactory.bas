Attribute VB_Name = "VtHttpaFactory"
Option Explicit

Public Function VtHttpaCreateService() As CVtHttpaService
    Set VtHttpaCreateService = New CVtHttpaService
End Function


Public Function VtHttpaCreateRequest() As CVtHttpaRequest

   Set VtHttpaCreateRequest = New CVtHttpaRequest
   

End Function


Public Function VtHttpaCreateResponse() As CVtHttpaResponse

   Set VtHttpaCreateResponse = New CVtHttpaResponse
   

End Function

