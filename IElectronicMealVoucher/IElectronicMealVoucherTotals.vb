Public Interface IElectronicMealVoucherTotals
    Function Totals(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode
    Function Check(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicMealVoucherReturnCode
End Interface
