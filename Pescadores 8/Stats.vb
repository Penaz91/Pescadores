Imports Microsoft.Office.Interop
Module Stats
    Function GCattFin()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattFin", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function GCattPerEsca()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattPerEsca", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function CMeseEsca()
        Dim obj As New Access.Application
        obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        obj.Visible = True
        obj.DoCmd.OpenReport("CMeseEsca", Access.AcView.acViewReport, , , Access.AcWindowMode.acWindowNormal, )
        Return "OK"
    End Function
    Function GProgCattMese()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GProgressioneCattMese", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function GcattMeteo()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattMeteo", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function RepEsche()
        Dim obj As New Access.Application
        obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        obj.Visible = True
        obj.DoCmd.OpenReport("Esche Registrate", Access.AcView.acViewReport, , , Access.AcWindowMode.acWindowNormal, )
        Return "OK"
    End Function
    Function RepPrede()
        Dim obj As New Access.Application
        obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        obj.Visible = True
        obj.DoCmd.OpenReport("Specie Catturate", Access.AcView.acViewReport, , , Access.AcWindowMode.acWindowNormal, )
        Return "OK"
    End Function
    Function RepFin()
        Dim obj As New Access.Application
        obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        obj.Visible = True
        obj.DoCmd.OpenReport("Finali Registrati", Access.AcView.acViewReport, , , Access.AcWindowMode.acWindowNormal, )
        Return "OK"
    End Function
    Function GcattFondo()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattFondo", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function GcattAcqua()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattAcqua", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
    Function GcattOra()
        Dim Obj As New Access.Application()
        Obj.OpenCurrentDatabase("C:\Pescadores\Stats.accdb", Exclusive:=True)
        Obj.Visible = True
        Obj.DoCmd.OpenForm("GCattOre", Access.AcFormView.acNormal, , , , , )
        Return "OK"
    End Function
End Module
