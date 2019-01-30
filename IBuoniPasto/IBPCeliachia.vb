Public Interface IBPCeliachia
    Function PaymentCeliachia(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IBPReturnCode
    Function StornoCeliachia(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IBPReturnCode
End Interface
