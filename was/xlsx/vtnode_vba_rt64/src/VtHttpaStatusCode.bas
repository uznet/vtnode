Attribute VB_Name = "VtHttpaStatusCode"

Option Explicit

Public Enum HTTPStatusCode

    ' Success (2xx)
    HTTPA_OK = 200
    HTTPA_CREATED = 201
    HTTPA_NO_CONTENT = 204

    ' Redirection (3xx)
    HTTPA_FOUND = 302
    HTTPA_NOT_MODIFIED = 304

    ' Client Error (4xx)
    HTTPA_BAD_REQUEST = 400
    HTTPA_UNAUTHORIZED = 401
    HTTPA_FORBIDDEN = 403
    HTTPA_NOT_FOUND = 404
    HTTPA_METHOD_NOT_ALLOWED = 405
    HTTPA_CONFLICT = 409
    HTTPA_PAYLOAD_TOO_LARGE = 413
    HTTPA_UNSUPPORTED_MEDIA_TYPE = 415

    ' Server Error (5xx)
    HTTPA_INTERNAL_SERVER_ERROR = 500
    HTTPA_NOT_IMPLEMENTED = 501
    HTTPA_SERVICE_UNAVAILABLE = 503

End Enum


