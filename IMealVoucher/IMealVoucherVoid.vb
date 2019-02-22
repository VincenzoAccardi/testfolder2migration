Public Interface IMealVoucherVoid
    Function Void(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IMealVoucherReturnCode
    Function Check(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IMealVoucherReturnCode
End Interface
