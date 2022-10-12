Option Strict Off
Option Explicit On
Imports System
Imports System.Text
Imports Microsoft.VisualBasic
Imports TPDotnet.Pos
Public Class TAX : Inherits TPDotnet.Pos.TAX
#Region "Properties"
    Public Overridable Property szTaxGroupRuleDescription() As String
        Get
            szTaxGroupRuleDescription = m.Fields_Value("szTaxGroupRuleDescription")
        End Get
        Set(ByVal Value As String)
            m.Fields_Value("szTaxGroupRuleDescription") = Value
        End Set
    End Property
#End Region

    Protected Overrides Sub DefineFields()
        MyBase.DefineFields()
        m.Append("szTaxGroupRuleDescription", DataField.FIELD_TYPES.FIELD_TYPE_STRING)

    End Sub
End Class
