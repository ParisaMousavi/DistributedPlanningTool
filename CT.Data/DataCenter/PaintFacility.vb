Imports System
Imports System.ComponentModel
Imports System.Reflection

Namespace DataCenter
    Public Enum PaintFacility
        <ComponentModel.Description("Lead or Proto")> LeadOrProto
        <ComponentModel.Description("Lead Plant Paint")> LeadPlantPaint
        <ComponentModel.Description("Proto Paint")> ProtoPaint
    End Enum

End Namespace
