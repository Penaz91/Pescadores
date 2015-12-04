Imports System.IO
Imports Microsoft.Office.Interop
Public Class Pescadores
    Dim Id As String
    Dim sit As Integer = 1
    Dim buonx As String
    Dim ora As Date
    Dim oraO As String
    Dim VT As String
    Dim side As Integer = 1
    Private Sub Pescadores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Catture'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_CattureTableAdapter.Fill(Me.TabelleDataSet.Registro_Catture)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Spots'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_SpotsTableAdapter.Fill(Me.TabelleDataSet.Registro_Spots)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Specie'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_SpecieTableAdapter.Fill(Me.TabelleDataSet.Registro_Specie)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Posizioni'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_PosizioniTableAdapter.Fill(Me.TabelleDataSet.Registro_Posizioni)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Meteo'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_MeteoTableAdapter.Fill(Me.TabelleDataSet.Registro_Meteo)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Tipi_di_luogo'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Tipi_di_luogoTableAdapter.Fill(Me.TabelleDataSet.Registro_Tipi_di_luogo)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Luoghi_di_pesca'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Luoghi_di_pescaTableAdapter.Fill(Me.TabelleDataSet.Registro_Luoghi_di_pesca)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Metodi_Presentazione'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Metodi_PresentazioneTableAdapter.Fill(Me.TabelleDataSet.Registro_Metodi_Presentazione)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Finali'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_FinaliTableAdapter.Fill(Me.TabelleDataSet.Registro_Finali)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Durata_Sessioni'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Durata_SessioniTableAdapter.Fill(Me.TabelleDataSet.Registro_Durata_Sessioni)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_condizioni_Fondale'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_condizioni_FondaleTableAdapter.Fill(Me.TabelleDataSet.Registro_condizioni_Fondale)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Condizioni_Acqua'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Condizioni_AcquaTableAdapter.Fill(Me.TabelleDataSet.Registro_Condizioni_Acqua)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Esche'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_EscheTableAdapter.Fill(Me.TabelleDataSet.Registro_Esche)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Pescatori'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_PescatoriTableAdapter.Fill(Me.TabelleDataSet.Registro_Pescatori)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Canne'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_CanneTableAdapter.Fill(Me.TabelleDataSet.Registro_Canne)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Specie'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_SpecieTableAdapter.Fill(Me.TabelleDataSet.Registro_Specie)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Posizioni'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_PosizioniTableAdapter.Fill(Me.TabelleDataSet.Registro_Posizioni)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Meteo'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_MeteoTableAdapter.Fill(Me.TabelleDataSet.Registro_Meteo)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Tipi_di_luogo'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Tipi_di_luogoTableAdapter.Fill(Me.TabelleDataSet.Registro_Tipi_di_luogo)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Luoghi_di_pesca'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Luoghi_di_pescaTableAdapter.Fill(Me.TabelleDataSet.Registro_Luoghi_di_pesca)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Metodi_Presentazione'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Metodi_PresentazioneTableAdapter.Fill(Me.TabelleDataSet.Registro_Metodi_Presentazione)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Finali'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_FinaliTableAdapter.Fill(Me.TabelleDataSet.Registro_Finali)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Durata_Sessioni'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Durata_SessioniTableAdapter.Fill(Me.TabelleDataSet.Registro_Durata_Sessioni)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_condizioni_Fondale'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_condizioni_FondaleTableAdapter.Fill(Me.TabelleDataSet.Registro_condizioni_Fondale)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Condizioni_Acqua'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_Condizioni_AcquaTableAdapter.Fill(Me.TabelleDataSet.Registro_Condizioni_Acqua)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Esche'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_EscheTableAdapter.Fill(Me.TabelleDataSet.Registro_Esche)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Pescatori'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_PescatoriTableAdapter.Fill(Me.TabelleDataSet.Registro_Pescatori)
        'TODO: questa riga di codice carica i dati nella tabella 'TabelleDataSet.Registro_Canne'. È possibile spostarla o rimuoverla se necessario.
        Me.Registro_CanneTableAdapter.Fill(Me.TabelleDataSet.Registro_Canne)
        Dim Regs As String
        Dim Qh As String
        Me.CenterToScreen()
        Id = My.Computer.FileSystem.ReadAllText("ID.pcf")
        If Id = "1" Then
            Do
                Id = InputBox("Inserisci il tuo nome o nickname", "Identificazione proprietario")
                My.Computer.FileSystem.WriteAllText("ID.Pcf", Id, False)
                Msg.Text = "Benvenuto in Pescadores 8, " & Id & ": seleziona un'operazione da uno dei pulsati della sidebar"
            Loop Until Id <> "" And Id <> "1"
        Else
            Saluto()
        End If
        Regs = My.Computer.FileSystem.ReadAllText("Registri.pcf")
        If Regs = "Si" Then
            TabControl.SelectTab(TabRegs)
        End If
        Qh = My.Computer.FileSystem.ReadAllText("Qh.pcf")
        If Qh = "Off" Then
            QuickHelp.Active = False
        Else
            QuickHelp.Active = True
        End If
        Dim Salva As Integer
        Salva = Val(My.Computer.FileSystem.ReadAllText("Salva.pcf"))
        If Salva = 0 Then
            ChkSalva.Checked = False
            MinSalva.Enabled = False
            SaveTimer.Enabled = False
        ElseIf Salva <> 0 Then
            ChkSalva.Checked = True
            MinSalva.Enabled = True
            MinSalva.Value = Salva
            SaveTimer.Enabled = True
            SaveTimer.Interval = (Salva * 60000)
        End If
    End Sub
    Function Saluto()
        ora = DateAndTime.TimeString
        oraO = DateAndTime.Hour(ora)
        If oraO >= 6 And oraO < 12 Then
            buonx = "Buon Giorno"
        ElseIf oraO >= 12 And oraO < 18 Then
            buonx = "Buon Pomeriggio"
        ElseIf (oraO >= 18 And oraO < 24) Or (oraO >= 0 And oraO < 6) Then
            buonx = "Buona sera"
        End If
        Msg.Text = buonx & ", " & Id & ", seleziona un'operazione da uno dei pulsanti della sidebar"
        Return "OK"
    End Function
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ora = DateAndTime.TimeString
        OraSistema.Text = CStr(ora)
        Saluto()
    End Sub
    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        TabControl.SelectTab(2)
    End Sub
    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        TabControl.SelectTab(4)
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        LoadInfo()
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        MyBase.Close()
    End Sub
    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        TabControl.SelectTab(3)
        LoadOpzioni()
    End Sub
    Function LoadOpzioni()
        Dim Nome As String
        Dim Regs As String
        Dim Qh As String
        Qh = My.Computer.FileSystem.ReadAllText("Qh.pcf")
        Nome = My.Computer.FileSystem.ReadAllText("ID.pcf")
        Regs = My.Computer.FileSystem.ReadAllText("Registri.pcf")
        If Regs = "Si" Then
            ChkRegs.Checked = True
        Else
            ChkRegs.Checked = False
        End If
        Txtnome.Text = CStr(Nome)
        If Qh = "On" Then
            ChkQH.Checked = True
        ElseIf Qh = "Off" Then
            ChkQH.Checked = False
        End If
        Return "OK"
    End Function
    Function LoadInfo()
        TabControl.SelectTab(1)
        Label19.Text = "Edizione concessa in licenza a: " & CStr(Id)
        Return "OK"
    End Function
    Private Sub Apply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Apply.Click
        Dim Id As String
        Dim Regs As String
        Id = Txtnome.Text
        If ChkRegs.Checked = True Then
            Regs = "Si"
        Else
            Regs = "No"
        End If
        My.Computer.FileSystem.WriteAllText("ID.pcf", Id, False)
        My.Computer.FileSystem.WriteAllText("Registri.pcf", Regs, False)
        If ChkQH.Checked = True Then
            My.Computer.FileSystem.WriteAllText("Qh.pcf", "On", False)
        Else
            My.Computer.FileSystem.WriteAllText("Qh.pcf", "Off", False)
        End If
        Dim salva As String
        salva = CStr(MinSalva.Value)
        If ChkSalva.Checked = True Then
            My.Computer.FileSystem.WriteAllText("Salva.pcf", salva, False)
        Else
            My.Computer.FileSystem.WriteAllText("Salva.pcf", "0", False)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Canne()
    End Sub
    Function Canne()
        TabControl.SelectTab(TabRegCanne)
        If BindingNavigatorCountItem.Text = "di 0" Then
            TxtLC.Enabled = False
            TXTMC.Enabled = False
            TxtModC.Enabled = False
            TxtNC.Enabled = False
            CmbSens.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub Pescadores8_onclose(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        Salvataggio()
        TmpTimer.Enabled = True
    End Sub

    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        TxtLC.Enabled = True
        TXTMC.Enabled = True
        TxtModC.Enabled = True
        TxtNC.Enabled = True
        CmbSens.Enabled = True
        TxtNC.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Pescatori()
    End Sub

    Private Sub BindingNavigatorAddNewItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem1.Click
        TxtNP.Enabled = True
        TxtNP.Focus()
    End Sub

    Private Sub BindingNavigatorAddNewItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem2.Click
        TxtNE.Enabled = True
        TXTIE.Enabled = True
        TxtNoteE.Enabled = True
        TxtNE.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Esche()
    End Sub
    Function Esche()
        TabControl.SelectTab(TabRegEsche)
        If BindingNavigatorCountItem2.Text = "di 0" Then
            TxtNE.Enabled = False
            TXTIE.Enabled = False
            TxtNoteE.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub BindingNavigatorAddNewItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem3.Click
        TxtCA.Enabled = True
        TxtCA.Focus()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Condacqua()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        CondFondo()
    End Sub

    Private Sub BindingNavigatorAddNewItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem4.Click
        TxtCF.Enabled = True
        TxtCF.Focus()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Duratas()
    End Sub

    Private Sub BindingNavigatorAddNewItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem5.Click
        TxtDS.Enabled = True
        TxtDS.Focus()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Finali()
    End Sub
    Function Finali()
        TabControl.SelectTab(TabRegFin)
        If BindingNavigatorCountItem6.Text = "di 0" Then
            TxtSF.Enabled = False
            TxtLF.Enabled = False
            TxtNF.Enabled = False
            CmbMP.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub BindingNavigatorAddNewItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem6.Click
        TxtSF.Enabled = True
        TxtLF.Enabled = True
        TxtNF.Enabled = True
        CmbMP.Enabled = True
        TxtNF.Focus()
    End Sub
    Private Sub CmdMp_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMP.GotFocus
        BtnModMP.Visible = True
    End Sub
    Private Sub CmdMp_noFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbMP.LostFocus
        BtnModMP.Visible = False
    End Sub
    Private Sub BtnModMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModMP.Click
        Me.RegistroFinaliBindingSource.EndEdit()
        Button8.Visible = True
        RegMP()
    End Sub
    Function RegMP()
        TabControl.SelectTab(TabRegMP)
        If BindingNavigatorCountItem7.Text = "di 0" Then
            TxtMetodo.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.RegistroCanneBindingSource.EndEdit()
    End Sub

    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.RegistroFinaliBindingSource.MoveLast()
        TabControl.SelectTab(TabRegFin)
        Button8.Visible = False
        Me.RegistroMetodiPresentazioneBindingSource.EndEdit()
    End Sub

    Private Sub BindingNavigatorAddNewItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem7.Click
        TxtMetodo.Enabled = True
        TxtMetodo.Focus()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        RegMP()
    End Sub
    Private Sub BtnModTL_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModTL.Click
        VT = "Luogo"
        TipiLuogo()
        Me.RegistroLuoghiDiPescaBindingSource.EndEdit()
        BtnIndietro1.Visible = True
    End Sub
    Function RegLuogo()
        TabControl.SelectTab(TabRegLuo)
        If BindingNavigatorCountItem8.Text = "di 0" Then
            TxtL.Enabled = False
            CmbTL.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        RegLuogo()
    End Sub
    Private Sub BindingNavigatorAddNewItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem8.Click
        TxtL.Enabled = True
        CmbTL.Enabled = True
        TxtL.Focus()
    End Sub

    Private Sub BtnIndietro1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnIndietro1.Click
        If VT = "Luogo" Then
            RegLuogo()
            RegistroLuoghiDiPescaBindingSource.MoveLast()
            BtnIndietro1.Visible = False
            Me.RegistroTipiDiLuogoBindingSource.EndEdit()
        ElseIf VT = "Posizione" Then
            Posizioni()
            RegistroPosizioniBindingSource.MoveLast()
            BtnIndietro1.Visible = False
            Me.RegistroTipiDiLuogoBindingSource.EndEdit()
        End If
    End Sub
    Function TipiLuogo()
        TabControl.SelectTab(TabRegTL)
        If BindingNavigatorCountItem9.Text = "di 0" Then
            TxtTL.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub BindingNavigatorAddNewItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem9.Click
        TxtTL.Enabled = True
        TxtTL.Focus()
    End Sub

    Private Sub CmbTL_Focus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTL.GotFocus
        BtnModTL.Visible = True
    End Sub
    Private Sub CmbTL_NoFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTL.LostFocus
        BtnModTL.Visible = False
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TipiLuogo()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Posizioni()
    End Sub
    Function Posizioni()
        TabControl.SelectTab(TabRegPos)
        If BindingNavigatorCountItem11.Text = "di 0" Then
            TxtPos.Enabled = False
            CmbTL2.Enabled = False
        End If
        Return "OK"
    End Function

    Private Sub BindingNavigatorAddNewItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem10.Click
        TxtCondm.Enabled = True
        TxtCondm.Focus()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Meteo()
    End Sub
    Function Meteo()
        TabControl.SelectTab(TabRegM)
        If BindingNavigatorCountItem10.Text = "di 0" Then
            TxtCondm.Enabled = False
        End If
        Return "OK"
    End Function

    Private Sub BindingNavigatorAddNewItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem11.Click
        TxtPos.Enabled = True
        CmbTL2.Enabled = True
        TxtPos.Focus()
    End Sub

    Private Sub CmbTL2_Focus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTL2.GotFocus
        BtnModTL2.Visible = True
    End Sub
    Private Sub CmbTL2_NoFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTL2.LostFocus
        BtnModTL2.Visible = False
    End Sub

    Private Sub BtnModTL2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModTL2.Click
        VT = "Posizione"
        TabControl.SelectTab(TabRegTL)
        BtnIndietro1.Visible = True
    End Sub

    Private Sub BindingNavigatorAddNewItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem12.Click
        TxtNS.Enabled = True
        TxtAA.Enabled = True
        TxtNote1.Enabled = True
        TxtNS.Focus()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Specie()
    End Sub
    Function Specie()
        TabControl.SelectTab(TabregSP)
        If BindingNavigatorCountItem12.Text = "di 0" Then
            TxtNS.Enabled = False
            TxtAA.Enabled = False
            TxtNote1.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub ComboBox3_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmbluo.GotFocus
        BtnModLUO.Visible = True
    End Sub
    Private Sub ComboBox3_NoHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmbluo.LostFocus
        BtnModLUO.Visible = False
    End Sub

    Private Sub BtnModLUO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModLUO.Click
        Me.RegistroSpotsBindingSource.EndEdit()
        RegLuogo()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        VaiASpots()
    End Sub

    Private Sub BindingNavigatorAddNewItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem13.Click
        CmbSpotPos.Enabled = True
        TxtNsp.Enabled = True
        CmbPP.Enabled = True
        Cmbluo.Enabled = True
        TxtNsp.Focus()
    End Sub
    Function Salvataggio()
        Timer1.Enabled = False
        ProgSalva.Visible = True
        ProgSalva.Value = 0
        Msg.Text = "Salvataggio:"
        Me.RegistroCanneBindingSource.EndEdit()
        Me.Registro_CanneTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 6
        Me.RegistroTipiDiLuogoBindingSource.EndEdit()
        Me.Registro_Tipi_di_luogoTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 12
        Me.RegistroCondizioniAcquaBindingSource.EndEdit()
        Me.Registro_Condizioni_AcquaTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 18
        Me.RegistroCondizioniFondaleBindingSource.EndEdit()
        Me.Registro_condizioni_FondaleTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 24
        Me.RegistroDurataSessioniBindingSource.EndEdit()
        Me.Registro_Durata_SessioniTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 30
        Me.RegistroEscheBindingSource.EndEdit()
        Me.Registro_EscheTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 36
        Me.RegistroFinaliBindingSource.EndEdit()
        Me.Registro_FinaliTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 42
        Me.RegistroLuoghiDiPescaBindingSource.EndEdit()
        Me.Registro_Luoghi_di_pescaTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 48
        Me.RegistroMeteoBindingSource.EndEdit()
        Me.Registro_MeteoTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 54
        Me.RegistroMetodiPresentazioneBindingSource.EndEdit()
        Me.Registro_Metodi_PresentazioneTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 60
        Me.RegistroPescatoriBindingSource.EndEdit()
        Me.Registro_PescatoriTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 66
        Me.RegistroPosizioniBindingSource.EndEdit()
        Me.Registro_PosizioniTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 72
        Me.RegistroSpecieBindingSource.EndEdit()
        Me.Registro_SpecieTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 78
        Me.RegistroSpotsBindingSource.EndEdit()
        Me.Registro_SpotsTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 84
        Me.RegistroCattureBindingSource.EndEdit()
        Me.Registro_CattureTableAdapter.Update(TabelleDataSet)
        ProgSalva.Value = 100
        TmpTimer.Enabled = True
        Return "OK"
    End Function
    Private Sub TmpTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TmpTimer.Tick
        Timer1.Enabled = True
        ProgSalva.Visible = False
        Saluto()
        TmpTimer.Enabled = False
    End Sub
    Private Sub ChkSalva_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSalva.CheckedChanged
        If ChkSalva.Checked = True Then
            MinSalva.Enabled = True
        Else
            MinSalva.Enabled = False
        End If
    End Sub
    Private Sub SaveTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveTimer.Tick
        Salvataggio()
        TmpTimer.Enabled = True
    End Sub
    Private Sub cmbCattC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TxtCattMC.Text = TxtResMC.Text
        TxtCattModC.Text = txtResModC.Text
        TxtCattSC.Text = txtResSC.Text
        TxtCattLC.Text = TxtResLC.Text
    End Sub
    Private Sub CmbCattEsca_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TxtCattLF.Text = TxtResLF.Text
        TxtCattSF.Text = TxtResSF.Text
    End Sub
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Catture()
    End Sub
    Private Sub BindingNavigatorAddNewItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem14.Click
        Panel1.Enabled = True
        ComboBox2.Focus()
    End Sub
    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        CattureSalva()
        Button17.Visible = True
        Pescatori()
    End Sub
    Function Pescatori()
        TabControl.SelectTab(TabRegPesc)
        If BindingNavigatorCountItem1.Text = "di 0" Then
            TxtNP.Enabled = False
        End If
        Return "OK"
    End Function
    Function CattureSalva()
        Me.RegistroCattureBindingSource.EndEdit()
        Return "OK"
    End Function
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroPescatoriBindingSource.EndEdit()
        Me.Registro_PescatoriTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button17.Visible = False
    End Sub
    Private Sub ComboBox2_Focus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.GotFocus
        PictureBox8.Visible = True
    End Sub
    Private Sub ComboBox2_noFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.LostFocus
        PictureBox8.Visible = False
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroSpotsBindingSource.EndEdit()
        Me.Registro_SpotsTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button18.Visible = False
    End Sub
    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        CattureSalva()
        Button18.Visible = True
        VaiASpots()
    End Sub
    Function VaiASpots()
        TabControl.SelectTab(TabRegSpot)
        If BindingNavigatorCountItem13.Text = "di 0" Then
            TxtNsp.Enabled = False
            CmbSpotPos.Enabled = False
            CmbPP.Enabled = False
            Cmbluo.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        CattureSalva()
        Button19.Visible = True
        CondAcqua()
    End Sub
    Function CondAcqua()
        TabControl.SelectTab(TabRegCondA)
        If BindingNavigatorCountItem3.Text = "di 0" Then
            TxtCA.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroCondizioniAcquaBindingSource.EndEdit()
        Me.Registro_Condizioni_AcquaTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button19.Visible = False
    End Sub
    Private Sub CmbCattSpot_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattSpot.GotFocus
        PictureBox9.Visible = True
    End Sub
    Private Sub CmbCattSpot_lostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattSpot.LostFocus
        PictureBox9.Visible = False
        TxtCattLuogo.Text = TxtResLuogo.Text
        TxtCattPress.Text = TxtResPress.Text
        TxtCattPos.Text = TxtSourcePos.Text
    End Sub
    Private Sub CmbCattCA_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattCA.GotFocus
        PictureBox10.Visible = True
    End Sub
    Private Sub CmbCattCA_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattCA.LostFocus
        PictureBox10.Visible = False
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroCondizioniFondaleBindingSource.EndEdit()
        Me.Registro_condizioni_FondaleTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button20.Visible = False
    End Sub
    Private Sub CmbCattCF_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattCF.GotFocus
        PictureBox13.Visible = True
    End Sub
    Private Sub CmbCattCF_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattCF.LostFocus
        PictureBox13.Visible = False
    End Sub
    Private Sub PictureBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click
        CattureSalva()
        Button20.Visible = True
        CondFondo()
    End Sub
    Function CondFondo()
        TabControl.SelectTab(TabRegCondf)
        If BindingNavigatorCountItem4.Text = "di 0" Then
            TxtCF.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub CmbCattDS_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattDS.GotFocus
        PictureBox11.Visible = True
    End Sub
    Private Sub CmbCattDS_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattDS.LostFocus
        PictureBox11.Visible = False
    End Sub
    Function Duratas()
        TabControl.SelectTab(TabRegDS)
        If BindingNavigatorCountItem5.Text = "di 0" Then
            TxtDS.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        CattureSalva()
        Button21.Visible = True
        Duratas()
    End Sub
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroDurataSessioniBindingSource.EndEdit()
        Me.Registro_Durata_SessioniTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button21.Visible = False
    End Sub
    Private Sub CmbCattMeteo_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattMeteo.GotFocus
        PictureBox12.Visible = True
    End Sub
    Private Sub CmbCattMeteo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattMeteo.LostFocus
        PictureBox12.Visible = False
    End Sub
    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        CattureSalva()
        Button22.visible = True
        Meteo()
    End Sub
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroMeteoBindingSource.EndEdit()
        Me.Registro_MeteoTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button22.Visible = False
    End Sub
    Private Sub cmbCattC_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCattC.GotFocus
        PictureBox14.Visible = True
    End Sub
    Private Sub cmbCattC_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCattC.LostFocus
        PictureBox14.Visible = False
        TxtCattLC.Text = TxtResLC.Text
        TxtCattMC.Text = TxtResMC.Text
        TxtCattModC.Text = txtResModC.Text
        TxtCattSC.Text = txtResSC.Text
    End Sub
    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        CattureSalva()
        Button23.Visible = True
        Canne()
    End Sub
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroCanneBindingSource.EndEdit()
        Me.Registro_CanneTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button23.Visible = False
    End Sub
    Private Sub CmbCattEsca_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattEsca.GotFocus
        PictureBox15.Visible = True
    End Sub
    Private Sub CmbCattEsca_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattEsca.LostFocus
        PictureBox15.Visible = False
    End Sub
    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        CattureSalva()
        Button24.Visible = True
        Esche()
    End Sub
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroEscheBindingSource.EndEdit()
        Me.Registro_EscheTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button24.Visible = False
    End Sub
    Private Sub CmbCattFin_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattFin.GotFocus
        PictureBox16.Visible = True
    End Sub
    Private Sub CmbCattFin_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCattFin.LostFocus
        PictureBox16.Visible = False
        TxtCattLF.Text = TxtResLF.Text
        TxtCattSF.Text = TxtResSF.Text
        TxtCattMP.Text = TxtResMP.Text
    End Sub
    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        CattureSalva()
        Button25.Visible = True
        Finali()
    End Sub
    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroFinaliBindingSource.EndEdit()
        Me.Registro_FinaliTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button25.Visible = False
    End Sub
    Private Sub ComboBox3_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.GotFocus
        PictureBox17.Visible = True
    End Sub
    Private Sub ComboBox3_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.LostFocus
        PictureBox17.Visible = False
    End Sub
    Private Sub PictureBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox17.Click
        CattureSalva()
        Button26.visible = True
        Specie()
    End Sub
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroSpecieBindingSource.EndEdit()
        Me.Registro_SpecieTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button26.Visible = False
    End Sub
    Private Sub CmbSpotPos_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSpotPos.GotFocus
        PictureBox18.Visible = True
    End Sub
    Private Sub CmbSpotPos_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSpotPos.LostFocus
        PictureBox18.Visible = False
    End Sub
    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        CattureSalva()
        Button27.Visible = True
        Posizioni()
    End Sub
    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Me.RegistroCattureBindingSource.MoveLast()
        Me.RegistroPosizioniBindingSource.EndEdit()
        Me.Registro_PosizioniTableAdapter.Update(TabelleDataSet)
        TabControl.SelectTab(TabRegCatt)
        Button27.Visible = False
    End Sub
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        GProgCattMese()
    End Sub
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        CMeseEsca()
    End Sub
    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        GCattPerEsca()
    End Sub
    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        GCattFin()
    End Sub
    Function Catture()
        TabControl.SelectTab(TabRegCatt)
        If BindingNavigatorCountItem14.Text = "di 0" Then
            Panel1.Enabled = False
        End If
        Return "OK"
    End Function
    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Catture()
    End Sub
    Private Sub RegistroCanneToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroCanneToolStripMenuItem.Click
        Canne()
    End Sub
    Private Sub RegistroPescatoriToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroPescatoriToolStripMenuItem.Click
        Pescatori()
    End Sub
    Private Sub RegistroEscheToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroEscheToolStripMenuItem.Click
        Esche()
    End Sub
    Private Sub RegistroCondizioniAcquaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroCondizioniAcquaToolStripMenuItem.Click
        CondAcqua()
    End Sub
    Private Sub RegistroCondizioniFondaleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroCondizioniFondaleToolStripMenuItem.Click
        CondFondo()
    End Sub
    Private Sub RegistroDurataSessioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroDurataSessioniToolStripMenuItem.Click
        Duratas()
    End Sub
    Private Sub RegistroFinaliToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroFinaliToolStripMenuItem.Click
        Finali()
    End Sub
    Private Sub RegistroMetodiDiPresentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroMetodiDiPresentToolStripMenuItem.Click
        RegMP()
    End Sub
    Private Sub RegistroLuoghiDiPescaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroLuoghiDiPescaToolStripMenuItem.Click
        RegLuogo()
    End Sub
    Private Sub RegistroTipiDiLuogoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroTipiDiLuogoToolStripMenuItem.Click
        TipiLuogo()
    End Sub
    Private Sub RegistroPosizioniToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroPosizioniToolStripMenuItem.Click
        Posizioni()
    End Sub
    Private Sub RegistroMeteoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroMeteoToolStripMenuItem.Click
        Meteo()
    End Sub
    Private Sub RegistroSpecieItticheToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroSpecieItticheToolStripMenuItem.Click
        Specie()
    End Sub
    Private Sub RegistroSpotsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistroSpotsToolStripMenuItem.Click
        VaiASpots()
    End Sub
    Private Sub MenRegs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenRegs.Click
        TabControl.SelectTab(2)
    End Sub
    Private Sub MenStats_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenStats.Click
        TabControl.SelectTab(4)
    End Sub
    Private Sub ProgressioneDelleCatturePerMesiDellannoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressioneDelleCatturePerMesiDellannoToolStripMenuItem.Click
        GProgCattMese()
    End Sub
    Private Sub CatturePerEscaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CatturePerEscaToolStripMenuItem.Click
        GCattPerEsca()
    End Sub
    Private Sub CattureRaggruppatePerEscaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CattureRaggruppatePerEscaToolStripMenuItem.Click
        CMeseEsca()
    End Sub
    Private Sub FinaliPiùEfficaciToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinaliPiùEfficaciToolStripMenuItem.Click
        GCattFin()
    End Sub
    Private Sub MenOpz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenOpz.Click
        TabControl.SelectTab(3)
        LoadOpzioni()
    End Sub
    Private Sub MenInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenInfo.Click
        LoadInfo()
    End Sub
    Private Sub MenEsci_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenEsci.Click
        MyBase.Close()
    End Sub
    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Dim Percorso As String
        Percorso = CStr(InputBox("Inserisci il percorso dove salvare il backup (deve terminare con \)"))
        If Percorso <> "" Then
            For Each Foundfile In My.Computer.FileSystem.GetFiles("C:\Pescadores\", FileIO.SearchOption.SearchTopLevelOnly, "Tabelle.accdb")
                My.Computer.FileSystem.CopyFile(Foundfile, Percorso & "Tabelle.accdb")
            Next
            For Each Foundfile In My.Computer.FileSystem.GetFiles(Percorso, FileIO.SearchOption.SearchTopLevelOnly, "Tabelle.accdb")
                MsgBox("Backup Riuscito")
            Next
        Else
            MsgBox("Backup Annullato")
        End If
    End Sub
    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        GcattMeteo()
    End Sub
    Private Sub CattureASecondaDelMeteoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CattureASecondaDelMeteoToolStripMenuItem.Click
        GcattMeteo()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        RepEsche()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        RepPrede()
    End Sub
    Private Sub EscheRegistrateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EscheRegistrateToolStripMenuItem.Click
        RepEsche()
    End Sub
    Private Sub PredeCatturateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PredeCatturateToolStripMenuItem.Click
        RepPrede()
    End Sub
    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        RepFin()
    End Sub
    Private Sub FinaliRegistratiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinaliRegistratiToolStripMenuItem.Click
        RepFin()
    End Sub
    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        GcattFondo()
    End Sub
    Private Sub CattureASecondaDelFondoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CattureASecondaDelFondoToolStripMenuItem.Click
        GcattFondo()
    End Sub
    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        GcattAcqua()
    End Sub
    Private Sub CattureASecondaDelleCondizioniDellacquaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CattureASecondaDelleCondizioniDellacquaToolStripMenuItem.Click
        GcattAcqua()
    End Sub
    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        GcattOra()
    End Sub
    Private Sub CattureASecondaDelloraDelGiornoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CattureASecondaDelloraDelGiornoToolStripMenuItem.Click
        GcattOra()
    End Sub
    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        Dim X As String
        X = CStr(InputBox("Le impostazioni di Pescadores saranno resettate. Scrivi YES per cancellare"))
        If X = "YES" Then
            My.Computer.FileSystem.WriteAllText("Id.pcf", "1", False)
            My.Computer.FileSystem.WriteAllText("Qh.pcf", "On", False)
            My.Computer.FileSystem.WriteAllText("Registri.pcf", "No", False)
            My.Computer.FileSystem.WriteAllText("Salva.pcf", "0", False)
            MsgBox("Il Programma verrà ora chiuso per applicare le nuove impostazioni")
            MyBase.Close()
        End If
    End Sub
    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("Copying.txt")
    End Sub
    Private Sub BtnCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub
    Private Sub BtnXc_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnXC.MouseHover
        BtnXC.Image = Pescadores_8.My.Resources.Expand_Collapset
    End Sub
    Private Sub BtnCollapse_NoHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnXC.MouseLeave
        BtnXC.Image = Pescadores_8.My.Resources.Expand_Collapse
    End Sub
    Private Sub BtnXc_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnXC.Click
        If Side = 1 Then
            PictureBox2.Visible = False
            PictureBox3.Visible = False
            PictureBox4.Visible = False
            PictureBox7.Visible = False
            PictureBox5.Visible = False
            PictureBox6.Visible = False
            BtnXC.Location = New Point(0, 100)
            OraSistema.Visible = False
            TabControl.Location = New Point(10, 125)
            TabControl.Size = New Size(1048, 600)
            Msg.Location = New Point(6, 103)
            ProgSalva.Location = New Point(105, 106)
            ProgSalva.Size = New Size(889, 13)
            side = 0
        Else
            If Side = 0 Then
                PictureBox2.Visible = True
                PictureBox3.Visible = True
                PictureBox4.Visible = True
                PictureBox7.Visible = True
                PictureBox5.Visible = True
                PictureBox6.Visible = True
                OraSistema.Visible = True
                BtnXC.Location = New Point(95, 100)
                TabControl.Location = New Point(106, 125)
                TabControl.Size = New Size(957, 600)
                Msg.Location = New Point(106, 103)
                ProgSalva.Location = New Point(205, 106)
                ProgSalva.Size = New Size(799, 13)
                side = 1
            End If
        End If
    End Sub
    Private Sub OraSistema_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OraSistema.Click
        DataOdierna.Visible = True
    End Sub
    Private Sub Tab_Hover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataOdierna.MouseLeave
        DataOdierna.Visible = False
    End Sub
End Class
