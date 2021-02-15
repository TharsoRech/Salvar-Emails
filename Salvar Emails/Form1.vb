Imports System.Net
Imports Ionic.Zip
Imports MetroFramework
Imports Microsoft.Exchange.WebServices.Data
Imports System.Text.RegularExpressions
Imports System.Text

Public Class Form1
    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        Me.Close()
    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        Try
            If login() = True Then
                MetroTabControl1.SelectedTab = TabPage2
                EnableTab(TabPage1, False)
                EnableTab(TabPage2, True)
            Else
                MetroMessageBox.Show(Me, "Não foi possivel Logar nessa conta,Verifique se a conta está certa", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function CheckedNames(theNodes As System.Windows.Forms.TreeNodeCollection) As List(Of [String])
        Dim aResult As New List(Of [String])()

        If theNodes IsNot Nothing Then
            For Each aNode As System.Windows.Forms.TreeNode In theNodes
                If aNode.Checked Then
                    aResult.Add(aNode.Text)
                End If

                aResult.AddRange(CheckedNames(aNode.Nodes))
            Next
        End If

        Return aResult
    End Function



    Public Shared Sub EnableTab(page As TabPage, enable As Boolean)
        For Each ctl As Control In page.Controls
            ctl.Enabled = enable
        Next
    End Sub

    Private Function FindAllSubFolders(ByVal service As ExchangeService, ByVal parentFolderId As FolderId, ByVal node As TreeNode) As String
        Try
            'search for sub folders
            Dim folderView As FolderView = New FolderView(1000)
            Dim foundFolders As FindFoldersResults = service.FindFolders(parentFolderId, folderView)

            ' Add the list to the growing complete list

            If foundFolders.TotalCount > 0 Then
            Else
                Return "no"
                Exit Function
            End If
            ' Now recurse
            For Each folder As Folder In foundFolders
                folder.Load()
                node.Nodes.Add(folder.DisplayName).Checked = True
                Dim subfolders As String = FindAllSubFolders2(service, folder.Id, node)
                If subfolders <> "no" Then
                    node.LastNode.Nodes.Add(subfolders).Checked = True
                    FindAllSubFolders(service, findfolders45(subfolders), node.LastNode.LastNode)
                End If
            Next
        Catch ex As Exception
            Return "no"
            ' EnviarReceber.Task1.adderror(ex.GetType.ToString, ex.Message, ex.StackTrace, Now.ToString)
        End Try
    End Function
    Private Function FindAllSubFolders2(ByVal service As ExchangeService, ByVal parentFolderId As FolderId, ByVal node As TreeNode) As String
        Try
            'search for sub folders
            Dim folderView As FolderView = New FolderView(1000)
            Dim foundFolders As FindFoldersResults = service.FindFolders(parentFolderId, folderView)

            ' Add the list to the growing complete list

            If foundFolders.TotalCount > 0 Then
            Else
                Return "no"
                Exit Function
            End If
            ' Now recurse
            For Each folder As Folder In foundFolders
                folder.Load()
                Return (folder.DisplayName)
            Next
        Catch ex As Exception
            Return "no"

        End Try
    End Function
    Public Function findfolders45(ByVal namefolder As String) As FolderId
        Try
            Dim user As String = Email.Text
            Dim pass As String = Senha.Text
            Dim server As String = Servidor.Text
            Dim service As New Global.Microsoft.Exchange.WebServices.Data.ExchangeService
            service.Credentials = New WebCredentials(user, pass, Environment.UserDomainName)
            service.Url = New Uri("https://" & server & "/ews/Exchange.asmx")





            Dim allFoldersType As ExtendedPropertyDefinition = New ExtendedPropertyDefinition(13825, MapiPropertyType.Integer)
            Dim rootFolderId As FolderId = New FolderId(WellKnownFolderName.Root)
            Dim folderView As FolderView = New FolderView(1000)
            folderView.Traversal = FolderTraversal.Deep
            Dim searchFilter2 As SearchFilter = New SearchFilter.IsEqualTo(FolderSchema.DisplayName, namefolder)
            Dim searchFilterCollection As SearchFilter.SearchFilterCollection = New SearchFilter.SearchFilterCollection(LogicalOperator.And)
            searchFilterCollection.Add(searchFilter2)

            Dim findFoldersResults As FindFoldersResults = service.FindFolders(rootFolderId, searchFilterCollection, folderView)

            If findFoldersResults.Folders.Count > 0 Then
                Dim allItemsFolder As Folder = findFoldersResults.Folders(0)
                Return allItemsFolder.Id
            End If


        Catch ex As ServiceRequestException



            '  MetroMessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As WebException


            '  MetroMessageBox.Show(Me, ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Function

    Public Function login() As Boolean
        Dim ok As Boolean = False
        Try
            Dim user As String = Email.Text
            Dim pass As String = Senha.Text
            Dim server As String = Servidor.Text
            Dim service As New Global.Microsoft.Exchange.WebServices.Data.ExchangeService
            service.Credentials = New WebCredentials(user, pass, Environment.UserDomainName)
            service.Url = New Uri("https://" & server & "/ews/Exchange.asmx")
            Dim rootfolder As Folder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot)
            'A GetFolder operation has been performed.
            'Now do something with the folder, such as display each child folder's name and id.
            rootfolder.Load()
            If rootfolder.FindFolders(New FolderView(100)).Count <> 0 Then
                ok = True
                Return ok
            End If
            Return ok
        Catch ex As Exception
            Return ok
        End Try
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            EnableTab(TabPage2, False)
            EnableTab(TabPage3, False)
            If My.Settings.ServerPadrao <> "." Then
                Servidor.Text = My.Settings.ServerPadrao
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Try
            If FolderBrowserDialog1.ShowDialog <> DialogResult.Abort Then
                Caminho.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        Try
            If IO.Directory.Exists(Caminho.Text) And SaveEmails.CheckState = CheckState.Checked Then
                EnableTab(TabPage2, False)
                EnableTab(TabPage3, True)
                MetroTabControl1.SelectedTab = TabPage3
                DataEscolhida.Text = Data.Value.ToShortDateString
                BackgroundWorker1.RunWorkerAsync()
            End If
            If IO.Directory.Exists(Caminho.Text) = False And SaveEmails.CheckState = CheckState.Checked Then
                MetroMessageBox.Show(Me, "Caminho para salvar emails Inexistente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            If SaveEmails.Checked = False And DeleteEmails.Checked = False Then
                MetroMessageBox.Show(Me, "Nenhuma ação selecionada,por favor selecione alguma antes de prosseguir", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            If SaveEmails.Checked = False And DeleteEmails.Checked = True Then
                EnableTab(TabPage2, False)
                EnableTab(TabPage3, True)
                MetroTabControl1.SelectedTab = TabPage3
                DataEscolhida.Text = Data.Value.ToShortDateString
                BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton6_Click(sender As Object, e As EventArgs) Handles MetroButton6.Click
        Try
            BackgroundWorker1.CancelAsync()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ProcurarEmails(ByVal user As String, ByVal pass As String, ByVal server As String, ByVal mailbox As String)

        If BackgroundWorker1.CancellationPending Then
            Exit Sub
        End If
        Dim folder1 As String = ""
        Dim service As New Global.Microsoft.Exchange.WebServices.Data.ExchangeService
        service.Credentials = New WebCredentials(user, pass, Environment.UserDomainName)
        service.Url = New Uri("https://" & server & "/ews/Exchange.asmx")
        service.PreAuthenticate = True
        service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.PrincipalName, mailbox)
        Try
            If BackgroundWorker1.CancellationPending Then
                Exit Sub
            End If
            Dim mb As Mailbox = New Mailbox(mailbox)
            Dim fid As FolderId = New FolderId(WellKnownFolderName.MsgFolderRoot, mb)
            Dim inbox As Folder = Folder.Bind(service, fid)
            Dim findResults As FindItemsResults(Of Item)
            Dim offset As Integer = 0
            Dim pageSize As Integer = 1000
            Dim more As Boolean = True
            Dim view As ItemView = New ItemView(pageSize)
            Me.Invoke(Sub() FolderProgress.Maximum = inbox.FindFolders(New FolderView(100)).Count)
            For Each folder As Folder In inbox.FindFolders(New FolderView(100))
                Me.Invoke(Sub() folder1 = folder.DisplayName)
                Me.Invoke(Sub() CurrentFolder.Text = folder.DisplayName)
                Me.Invoke(Sub() FolderProgress.Value += 1)
                findResults = folder.FindItems(New ItemView(1000))
                Me.Invoke(Sub() EmailsProgress.Value = 0)
                Me.Invoke(Sub() EmailsProgress.Maximum = findResults.Count)
                If BackgroundWorker1.CancellationPending Then
                    Exit Sub
                End If
                For Each x10 As Item In findResults.Items

                    If BackgroundWorker1.CancellationPending Then
                        Exit Sub
                    End If
                    Me.Invoke(Sub() EmailsProgress.Value += 1)
                    Dim value As String = ""
                    If x10.GetType() Is GetType(EmailMessage) Then
                        Dim d As EmailMessage = TryCast(x10, EmailMessage)
                        If (d.DateTimeReceived - Data.Value).TotalDays < 0 Then
                            d.Load()

                            If ChecarPalavraschave.Checked = True Then
                                Dim timereceived As Date = d.DateTimeReceived
                                Dim fromadress As String = "Remententenãoidentificado"
                                If d.From IsNot Nothing Then
                                    fromadress = d.From.Address
                                End If
                                Dim subjectx As String = "SemAssunto"
                                Dim Bodyx As String = "SemCorpoDeEmail"
                                subjectx = d.Subject
                                Bodyx = d.Body.Text
                                Dim Props As PropertySet = New PropertySet(BasePropertySet.IdOnly)
                                Props.Add(EmailMessageSchema.Subject)
                                Props.Add(EmailMessageSchema.Body)
                                d.Load(Props)

                                For Each x6 As TreeNode In TreeView2.Nodes
                                    If Checkaloneword(subjectx, x6.Text) = True Or Checkaloneword(Bodyx, x6.Text) = True Then
                                        Me.Invoke(Sub() addchave(x6.Text, subjectx & Space(2) & "da caixa de entrada do usuário:" & Space(2) & mailbox & Space(2) & "na pasta:" & folder1))
                                        If SaveEmails.Checked = True Then
                                            Me.Invoke(Sub() CurrentEmail.Text = "Salvando Mensagem" & Space(2) & subjectx)
                                            value = timereceived
                                            Dim Res As String = ""
                                            For Each c As Char In value
                                                If IsNumeric(c) Then
                                                    Res = Res & c
                                                End If
                                            Next
                                            Res = Res & timereceived.Second & fromadress
                                            d.Load(New PropertySet(ItemSchema.MimeContent))
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            Dim mc As MimeContent = d.MimeContent
                                            If IO.Directory.Exists(Caminho.Text & "\" & folder1) = False Then
                                                IO.Directory.CreateDirectory(Caminho.Text & "\" & mailbox & "\" & folder1)
                                            End If
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            Dim fs As IO.FileStream = New IO.FileStream(Caminho.Text & "\" & mailbox & "\" & folder1 & "\" & Res & ".eml", IO.FileMode.Create)
                                            fs.Write(mc.Content, 0, mc.Content.Length)
                                            fs.Close()
                                            value = ""
                                            Me.Invoke(Sub() CurrentEmail.Text = "O email de assunto" & Space(2) & subjectx & Space(2) & " foi salvo com sucesso no arquivo" & Space(2) & Res & ".eml" & Space(2) & "Na pasta" & Space(2) & folder1)
                                            Res = ""

                                        End If
                                        If BackgroundWorker1.CancellationPending Then
                                            Exit Sub
                                        End If
                                        If DeleteEmails.Checked = True Then
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            d.Load()
                                            Me.Invoke(Sub() CurrentEmail.Text = "Deletando Mensagem" & Space(2) & subjectx)
                                            d.Delete(DeleteMode.HardDelete)
                                            Me.Invoke(Sub() CurrentEmail.Text = "A seguinte mensagem foi deletada com sucesso do servidor:" & Space(2) & d.Subject)
                                        End If
                                    End If
                                Next
                            Else
                                Dim timereceived As Date = d.DateTimeReceived
                                Dim fromadress As String = "Remententenãoidentificado"
                                If d.From IsNot Nothing Then
                                    fromadress = d.From.Address
                                End If
                                Dim subjectx As String = "SemAssunto"
                                Dim Bodyx As String = "SemCorpoDeEmail"
                                subjectx = d.Subject
                                Bodyx = d.Body.Text
                                Dim Props As PropertySet = New PropertySet(BasePropertySet.IdOnly)
                                Props.Add(EmailMessageSchema.Subject)
                                Props.Add(EmailMessageSchema.Body)
                                d.Load(Props)

                                If SaveEmails.Checked = True Then
                                    Me.Invoke(Sub() CurrentEmail.Text = "Salvando Mensagem" & Space(2) & subjectx)
                                    value = timereceived.Date.ToString
                                    Dim Res As String = ""
                                    For Each c As Char In value
                                        If IsNumeric(c) Then
                                            Res = Res & c
                                        End If
                                    Next
                                    Res = Res & timereceived.Second & fromadress
                                    d.Load(New PropertySet(ItemSchema.MimeContent))
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    Dim mc As MimeContent = d.MimeContent
                                    If IO.Directory.Exists(Caminho.Text & "\" & folder1) = False Then
                                        IO.Directory.CreateDirectory(Caminho.Text & "\" & mailbox & "\" & folder1)
                                    End If
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    Dim fs As IO.FileStream = New IO.FileStream(Caminho.Text & "\" & mailbox & "\" & folder1 & "\" & Res & ".eml", IO.FileMode.Create)
                                    fs.Write(mc.Content, 0, mc.Content.Length)
                                    fs.Close()
                                    value = ""
                                    Me.Invoke(Sub() CurrentEmail.Text = "O email de assunto" & Space(2) & subjectx & Space(2) & " foi salvo com sucesso no arquivo" & Space(2) & Res & ".eml" & Space(2) & "Na pasta" & Space(2) & folder1)

                                    Res = ""

                                End If
                                If BackgroundWorker1.CancellationPending Then
                                    Exit Sub
                                End If
                                If DeleteEmails.Checked = True Then
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    d.Load()
                                    Me.Invoke(Sub() CurrentEmail.Text = "Deletando Mensagem" & Space(2) & subjectx)
                                    d.Delete(DeleteMode.HardDelete)
                                    Me.Invoke(Sub() CurrentEmail.Text = "A seguinte mensagem foi deletada com sucesso do servidor:" & Space(2) & subjectx)
                                End If
                            End If
                        End If
                    End If
                Next
            Next
            If findResults.TotalCount > 1000 Then
                If BackgroundWorker1.CancellationPending Then
                    Exit Sub
                End If
                ProcurarEmails(user, pass, server, folder1, findResults.NextPageOffset)
            End If

        Catch ex As Exception
            Me.Invoke(Sub() adderror(ex.GetType.ToString, ex.Message, ex.StackTrace, Now.ToString))
            '  adderror(ex.get)
            '  MetroMessageBox.Show(Me, ex.Message & ex.StackTrace, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Shared Function Checkaloneword(ByVal email As String, ByVal SB As String) As Boolean
        Dim padraoRegex As String = "\b(" & SB & ")\b"
        Dim verifica As New RegularExpressions.Regex(padraoRegex, RegexOptions.IgnoreCase)
        Dim valida As Boolean = False
        'verifica se foi informado um email
        If String.IsNullOrEmpty(email) Then
            valida = False
        Else
            'usar IsMatch para validar o email
            valida = verifica.IsMatch(email)
        End If
        'retorna o valor
        Return valida
    End Function
    Public Sub ProcurarEmails(ByVal user As String, ByVal pass As String, ByVal server As String, ByVal mailbox As String, ByVal offset As Integer)

        If BackgroundWorker1.CancellationPending Then
            Exit Sub
        End If
        Dim folder1 As String = ""
        Dim service As New Global.Microsoft.Exchange.WebServices.Data.ExchangeService
        service.Credentials = New WebCredentials(user, pass, Environment.UserDomainName)
        service.Url = New Uri("https://" & server & "/ews/Exchange.asmx")
        service.PreAuthenticate = True
        service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.PrincipalName, mailbox)
        Try
            If BackgroundWorker1.CancellationPending Then
                Exit Sub
            End If
            Dim mb As Mailbox = New Mailbox(mailbox)
            Dim fid As FolderId = New FolderId(WellKnownFolderName.MsgFolderRoot, mb)
            Dim inbox As Folder = Folder.Bind(service, fid)
            Dim findResults As FindItemsResults(Of Item)
            Dim pageSize As Integer = 1000
            Dim more As Boolean = True
            Dim view As ItemView = New ItemView(pageSize, offset)
            Me.Invoke(Sub() FolderProgress.Maximum = inbox.FindFolders(New FolderView(100)).Count)
            For Each folder As Folder In inbox.FindFolders(New FolderView(100))
                Me.Invoke(Sub() folder1 = folder.DisplayName)
                Me.Invoke(Sub() CurrentFolder.Text = folder.DisplayName)
                Me.Invoke(Sub() FolderProgress.Value += 1)
                findResults = folder.FindItems(New ItemView(1000))
                Me.Invoke(Sub() EmailsProgress.Value = 0)
                Me.Invoke(Sub() EmailsProgress.Maximum = findResults.Count)
                If BackgroundWorker1.CancellationPending Then
                    Exit Sub
                End If
                For Each x10 As Item In findResults.Items

                    If BackgroundWorker1.CancellationPending Then
                        Exit Sub
                    End If
                    Me.Invoke(Sub() EmailsProgress.Value += 1)
                    Dim value As String = ""
                    If x10.GetType() Is GetType(EmailMessage) Then
                        Dim d As EmailMessage = TryCast(x10, EmailMessage)
                        If (d.DateTimeReceived - Data.Value).TotalDays < 0 Then
                            d.Load()
                            If ChecarPalavraschave.Checked = True Then
                                Dim timereceived As Date = d.DateTimeReceived
                                Dim fromadress As String = "Remententenãoidentificado"
                                If d.From IsNot Nothing Then
                                    fromadress = d.From.Address
                                End If
                                Dim subjectx As String = "SemAssunto"
                                Dim Bodyx As String = "SemCorpoDeEmail"
                                Dim Props As PropertySet = New PropertySet(BasePropertySet.IdOnly)
                                Props.Add(EmailMessageSchema.Subject)
                                Props.Add(EmailMessageSchema.Body)
                                d.Load(Props)
                                subjectx = d.Subject
                                Bodyx = d.Body.Text
                                For Each x6 As TreeNode In TreeView2.Nodes
                                    If Checkaloneword(subjectx, x6.Text) = True Or Checkaloneword(Bodyx, x6.Text) = True Then
                                        Me.Invoke(Sub() addchave(x6.Text, subjectx & Space(2) & "da caixa de entrada do usuário:" & Space(2) & mailbox & Space(2) & "na pasta:" & folder1))
                                        If SaveEmails.Checked = True Then
                                            Me.Invoke(Sub() CurrentEmail.Text = "Salvando Mensagem" & Space(2) & subjectx)
                                            value = timereceived
                                            Dim Res As String = ""
                                            For Each c As Char In value
                                                If IsNumeric(c) Then
                                                    Res = Res & c
                                                End If
                                            Next
                                            Res = Res & timereceived.Second & fromadress
                                            d.Load(New PropertySet(ItemSchema.MimeContent))
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            Dim mc As MimeContent = d.MimeContent
                                            If IO.Directory.Exists(Caminho.Text & "\" & folder1) = False Then
                                                IO.Directory.CreateDirectory(Caminho.Text & "\" & mailbox & "\" & folder1)
                                            End If
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            Dim fs As IO.FileStream = New IO.FileStream(Caminho.Text & "\" & mailbox & "\" & folder1 & "\" & Res & ".eml", IO.FileMode.Create)
                                            fs.Write(mc.Content, 0, mc.Content.Length)
                                            fs.Close()
                                            value = ""
                                            Me.Invoke(Sub() CurrentEmail.Text = "O email de assunto" & Space(2) & subjectx & Space(2) & " foi salvo com sucesso no arquivo" & Space(2) & Res & ".eml" & Space(2) & "Na pasta" & Space(2) & folder1)
                                            Res = ""

                                        End If
                                        If BackgroundWorker1.CancellationPending Then
                                            Exit Sub
                                        End If
                                        If DeleteEmails.Checked = True Then
                                            If BackgroundWorker1.CancellationPending Then
                                                Exit Sub
                                            End If
                                            d.Load()
                                            Me.Invoke(Sub() CurrentEmail.Text = "Deletando Mensagem" & Space(2) & subjectx)
                                            d.Delete(DeleteMode.HardDelete)
                                            Me.Invoke(Sub() CurrentEmail.Text = "A seguinte mensagem foi deletada com sucesso do servidor:" & Space(2) & d.Subject)
                                        End If
                                    End If
                                Next
                            Else
                                Dim timereceived As Date = d.DateTimeReceived
                                Dim fromadress As String = "Remententenãoidentificado"
                                If d.From IsNot Nothing Then
                                    fromadress = d.From.Address
                                End If
                                Dim subjectx As String = "SemAssunto"
                                Dim Bodyx As String = "SemCorpoDeEmail"
                                Dim Props As PropertySet = New PropertySet(BasePropertySet.IdOnly)
                                Props.Add(EmailMessageSchema.Subject)
                                Props.Add(EmailMessageSchema.Body)
                                d.Load(Props)
                                subjectx = d.Subject
                                Bodyx = d.Body.Text
                                If SaveEmails.Checked = True Then
                                    Me.Invoke(Sub() CurrentEmail.Text = "Salvando Mensagem" & Space(2) & subjectx)
                                    value = timereceived.Date.ToString
                                    Dim Res As String = ""
                                    For Each c As Char In value
                                        If IsNumeric(c) Then
                                            Res = Res & c
                                        End If
                                    Next
                                    Res = Res & timereceived.Second & fromadress
                                    d.Load(New PropertySet(ItemSchema.MimeContent))
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    Dim mc As MimeContent = d.MimeContent
                                    If IO.Directory.Exists(Caminho.Text & "\" & folder1) = False Then
                                        IO.Directory.CreateDirectory(Caminho.Text & "\" & mailbox & "\" & folder1)
                                    End If
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    Dim fs As IO.FileStream = New IO.FileStream(Caminho.Text & "\" & mailbox & "\" & folder1 & "\" & Res & ".eml", IO.FileMode.Create)
                                    fs.Write(mc.Content, 0, mc.Content.Length)
                                    fs.Close()
                                    value = ""
                                    Me.Invoke(Sub() CurrentEmail.Text = "O email de assunto" & Space(2) & subjectx & Space(2) & " foi salvo com sucesso no arquivo" & Space(2) & Res & ".eml" & Space(2) & "Na pasta" & Space(2) & folder1)

                                    Res = ""

                                End If
                                If BackgroundWorker1.CancellationPending Then
                                    Exit Sub
                                End If
                                If DeleteEmails.Checked = True Then
                                    If BackgroundWorker1.CancellationPending Then
                                        Exit Sub
                                    End If
                                    d.Load()
                                    Me.Invoke(Sub() CurrentEmail.Text = "Deletando Mensagem" & Space(2) & subjectx)
                                    d.Delete(DeleteMode.HardDelete)
                                    Me.Invoke(Sub() CurrentEmail.Text = "A seguinte mensagem foi deletada com sucesso do servidor:" & Space(2) & subjectx)
                                End If
                            End If
                        End If
                    End If
                Next
            Next
            If findResults.TotalCount > 1000 Then
                If BackgroundWorker1.CancellationPending Then
                    Exit Sub
                End If
                ProcurarEmails(user, pass, server, folder1, findResults.NextPageOffset)
            End If

        Catch ex As Exception
            Me.Invoke(Sub() adderror(ex.GetType.ToString, ex.Message, ex.StackTrace, Now.ToString))
            '  adderror(ex.get)
            '  MetroMessageBox.Show(Me, ex.Message & ex.StackTrace, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.Invoke(Sub() MailBoxProgress.Maximum = CheckedNames(TreeView1.Nodes).Count)
            Me.Invoke(Sub() FolderProgress.Value = 0)
            Me.Invoke(Sub() MailBoxProgress.Value = 0)
            Me.Invoke(Sub() EmailsProgress.Value = 0)
            Me.Invoke(Sub() CurrentFolder.Text = "")
            Me.Invoke(Sub() CurrentMailBox.Text = "")
            Me.Invoke(Sub() CurrentEmail.Text = "")
            For Each x1 As String In CheckedNames(TreeView1.Nodes)
                If BackgroundWorker1.CancellationPending Then
                    Exit Sub
                End If
                Me.Invoke(Sub() MailBoxProgress.Value += 1)
                Me.Invoke(Sub() CurrentMailBox.Text = x1)
                ProcurarEmails(Email.Text, Senha.Text, Servidor.Text, x1)
                If IO.Directory.Exists(Caminho.Text & "\" & x1) Then
                    Dim file As ZipFile = New ZipFile
                    file.AddDirectory(Caminho.Text & "\" & x1)
                    file.Save(Caminho.Text & "\" & x1 & ".zip")
                    IO.Directory.Delete(Caminho.Text & "\" & x1, True)
                    '   MetroMessageBox.Show(Me, "Emails Salvos com sucesso em:" & Caminho.Text & "\" & Email.Text, "Emails Salvos", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Next



        Catch ex As Exception
            adderror(ex.GetType.ToString, ex.Message, ex.StackTrace, Now.ToString)
        End Try
    End Sub
    Public Sub donework() Handles BackgroundWorker1.RunWorkerCompleted
        Try
            EnableTab(TabPage1, True)
            FolderProgress.Value = 0
            MailBoxProgress.Value = 0
            EmailsProgress.Value = 0
            CurrentFolder.Text = ""
            CurrentMailBox.Text = ""
            CurrentEmail.Text = ""
            ' EnableTab(TabPage3, False)
            ' MetroGrid1.Rows.Clear()
            ' MetroGrid2.Rows.Clear()
            MetroTabControl1.SelectedTab = TabPage1
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        Try
            My.Settings.ServerPadrao = Servidor.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton7_Click(sender As Object, e As EventArgs) Handles MetroButton7.Click
        Try
            TreeView2.Nodes.Add(PalavraChave.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroContextMenu1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MetroContextMenu1.Opening

    End Sub

    Private Sub RemoverPalavraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoverPalavraToolStripMenuItem.Click
        Try
            TreeView2.Nodes.Remove(TreeView2.SelectedNode)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton8_Click(sender As Object, e As EventArgs) Handles MetroButton8.Click
        Try
            OpenFD.Title = "Open a Text File"
            OpenFD.Filter = "Text Files|*.txt"
            If OpenFD.ShowDialog <> DialogResult.Abort Then
                Dim lines() As String = IO.File.ReadAllLines(OpenFD.FileName)
                For Each line As String In lines
                    TreeView2.Nodes.Add(line)
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub adderror(ByVal typeerror As String, ByVal errorr As String, ByVal stacktrace As String, ByVal datatime As String)
        Try

            ' Dim row As DataGridViewRow = DirectCast(MetroGrid1.Rows(0).Clone(), DataGridViewRow)
            Dim row1 As String() = New String() {typeerror, errorr, stacktrace, datatime}
            MetroGrid1.Rows.Add(row1)
            Me.Invoke(Sub() CountErrors.Text = "Número de erros achados:" & MetroGrid1.Rows.Count)
            MetroGrid1.Update()
            MetroGrid1.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub addchave(ByVal typeerror As String, ByVal errorr As String)
        Try

            ' Dim row As DataGridViewRow = DirectCast(MetroGrid1.Rows(0).Clone(), DataGridViewRow)
            Dim row1 As String() = New String() {typeerror, errorr}
            MetroGrid2.Rows.Add(row1)
            Me.Invoke(Sub() CountChaves.Text = "Número de chaves achadas:" & MetroGrid2.Rows.Count)
            MetroGrid2.Update()
            MetroGrid2.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton9_Click(sender As Object, e As EventArgs)
        Try
            Dim service As New Global.Microsoft.Exchange.WebServices.Data.ExchangeService
            service.Credentials = New WebCredentials("tharso.curia@bepo.com.br", "Garfo2014")
            service.PreAuthenticate = True
            service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.PrincipalName, "julio.capelete@bepo.com.br")
            service.Url = New Uri("https://outlook.office365.com/ews/Exchange.asmx")
            Dim mb As Mailbox = New Mailbox("julio.capelete@bepo.com.br")
            Dim fid As FolderId = New FolderId(WellKnownFolderName.Inbox, mb)
            Dim inbox As Folder = Folder.Bind(service, fid)
            'load items from mailbox inbox folder
            If (Not (inbox) Is Nothing) Then
                Dim items As FindItemsResults(Of Item) = inbox.FindItems(New ItemView(100))
                For Each item In items
                    item.Load()
                    MsgBox(item.Subject)
                Next
            End If



        Catch ex As Exception
            MsgBox(ex.Message & ex.StackTrace)
        End Try
    End Sub

    Private Sub MetroButton10_Click(sender As Object, e As EventArgs) Handles MetroButton10.Click
        Try
            OpenFD.Title = "Open a Text File"
            OpenFD.Filter = "Text Files|*.txt"
            If OpenFD.ShowDialog <> DialogResult.Abort Then
                Dim lines() As String = IO.File.ReadAllLines(OpenFD.FileName)
                For Each line As String In lines
                    TreeView1.Nodes.Add(line).Checked = True
                Next
                TreeView1.Sort()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton11_Click(sender As Object, e As EventArgs) Handles MetroButton11.Click
        Try
            TreeView1.Nodes.Add(MetroTextBox1.Text).Checked = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Try
            TreeView1.Nodes.Remove(TreeView1.SelectedNode)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton9_Click_1(sender As Object, e As EventArgs) Handles MetroButton9.Click
        Try
            If TreeView1.Nodes IsNot Nothing Then
                For Each aNode As System.Windows.Forms.TreeNode In TreeView1.Nodes
                    If aNode.Checked Then
                        TreeView1.Nodes.Remove(aNode)
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub MetroButton12_Click(sender As Object, e As EventArgs) Handles MetroButton12.Click
        Try
            TreeView2.Nodes.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton13_Click(sender As Object, e As EventArgs) Handles MetroButton13.Click
        Try
            For Each x1 As TreeNode In TreeView1.Nodes
                If x1.Checked = True Then
                    x1.Checked = False
                Else
                    x1.Checked = True
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton15_Click(sender As Object, e As EventArgs) Handles MetroButton15.Click
        Try
            MetroGrid1.Rows.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton14_Click(sender As Object, e As EventArgs) Handles MetroButton14.Click
        Try
            MetroGrid2.Rows.Clear()
        Catch ex As Exception

        End Try
    End Sub
End Class
