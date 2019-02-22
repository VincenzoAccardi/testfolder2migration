Public Interface IElectronicMealVoucherVoid
    Function Void(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode
    Function Check(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode
End Interface