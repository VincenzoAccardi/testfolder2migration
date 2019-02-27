Public Interface IElectronicFundsTransferClosure
    Function Closure(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode
    Function Check(ByRef Parameters As System.Collections.Generic.Dictionary(Of String, Object)) As IElectronicFundsTransferReturnCode
End Interface
