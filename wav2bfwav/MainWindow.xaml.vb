Imports System.IO

Class MainWindow
    Private inputPaths As New List(Of String)
    Private selectedFormat As String = "--dspadpcm"
    Private selectedFormatName As String = "dspadpcm"
    Private _tempExePath As String
    Private _tempDllPaths As New List(Of String)
    Private isProcessing As Boolean = False
    Private isSelecting As Boolean = False

    Public Class NaturalStringComparer
        Implements IComparer(Of String)

        Private Declare Unicode Function StrCmpLogicalW Lib "shlwapi.dll" (ByVal s1 As String, ByVal s2 As String) As Integer

        Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
            Return StrCmpLogicalW(x, y)
        End Function
    End Class

    Public Class FileItem
        Public Property FullPath As String
        Public ReadOnly Property FileName As String
            Get
                Return System.IO.Path.GetFileName(FullPath)
            End Get
        End Property

        Public Sub New(path As String)
            FullPath = path
        End Sub

        Public Overrides Function ToString() As String
            Return FileName
        End Function
    End Class

    Public Sub New()
        InitializeComponent()
        InitializeEmbeddedResources()
    End Sub

    Private Sub InitializeEmbeddedResources()
        Dim tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "wav2bfwav_temp")
        System.IO.Directory.CreateDirectory(tempDir)

        _tempExePath = System.IO.Path.Combine(tempDir, "NW4F_WaveConverter.exe")
        ExtractEmbeddedResource("NW4F_WaveConverter.exe", _tempExePath)

        Dim dllFiles = {
            "SoundFoundation.dll",
            "SoundFoundationCafe.dll",
            "ToolDevelopmentKit.dll",
            "WaveCodecCafe.dll"
        }

        For Each dllFile In dllFiles
            Dim dllPath = System.IO.Path.Combine(tempDir, dllFile)
            ExtractEmbeddedResource(dllFile, dllPath)
            _tempDllPaths.Add(dllPath)
        Next
    End Sub

    Private Sub ExtractEmbeddedResource(resourceName As String, outputPath As String)
        If Not System.IO.File.Exists(outputPath) Then
            Dim executingAssembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames = executingAssembly.GetManifestResourceNames()
            Dim fullResourceName = resourceNames.FirstOrDefault(Function(r) r.EndsWith(resourceName))

            If fullResourceName Is Nothing Then
                Dispatcher.Invoke(Sub() MessageBox.Show($"嵌入的资源未找到:{resourceName}", "错误", MessageBoxButton.OK, MessageBoxImage.Error))
                Return
            End If

            Using stream = executingAssembly.GetManifestResourceStream(fullResourceName)
                If stream Is Nothing Then
                    Dispatcher.Invoke(Sub() MessageBox.Show($"无法读取嵌入的资源:{resourceName}", "错误", MessageBoxButton.OK, MessageBoxImage.Error))
                    Return
                End If

                Dim buffer = New Byte(stream.Length - 1) {}
                stream.Read(buffer, 0, buffer.Length)
                System.IO.File.WriteAllBytes(outputPath, buffer)
            End Using
        End If
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        If Not System.IO.File.Exists(_tempExePath) Then
            MessageBox.Show("无法提取NW4F_WaveConverter.exe,请确保资源已正确嵌入.", "错误", MessageBoxButton.OK, MessageBoxImage.Error)
            btnProcess.IsEnabled = False
        End If
    End Sub

    Private Sub btnSelectFiles_Click(sender As Object, e As RoutedEventArgs)
        If isSelecting Then Return
        isSelecting = True

        Try
            If isProcessing Then
                MessageBox.Show("正在处理中,请等待当前任务完成", "提示", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            Dim openFileDialog As New Microsoft.Win32.OpenFileDialog()
            openFileDialog.Filter = "WAV文件(*.wav)|*.wav|所有文件(*.*)|*.*"
            openFileDialog.FilterIndex = 1
            openFileDialog.Multiselect = True
            openFileDialog.Title = "选择WAV文件"

            Dim result = openFileDialog.ShowDialog()
            If result = True Then
                Dim sortedFiles = openFileDialog.FileNames.OrderBy(Function(f) f, New NaturalStringComparer()).ToList()
                AddFiles(sortedFiles)
            End If
        Finally
            isSelecting = False
        End Try
    End Sub

    Private Sub btnSelectFolder_Click(sender As Object, e As RoutedEventArgs)
        If isSelecting Then Return
        isSelecting = True

        Try
            If isProcessing Then
                MessageBox.Show("正在处理中,请等待当前任务完成", "提示", MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            Dim folderDialog As New Microsoft.Win32.OpenFolderDialog()
            folderDialog.Title = "选择包含WAV文件的文件夹"

            Dim result = folderDialog.ShowDialog()
            If result = True Then
                Dim wavFiles = System.IO.Directory.GetFiles(folderDialog.FolderName, "*.wav", System.IO.SearchOption.AllDirectories)
                If wavFiles.Length > 0 Then
                    Dim sortedFiles = wavFiles.OrderBy(Function(f) f, New NaturalStringComparer()).ToList()
                    AddFiles(sortedFiles)
                Else
                    MessageBox.Show("该文件夹中没有找到WAV文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information)
                End If
            End If
        Finally
            isSelecting = False
        End Try
    End Sub

    Private Sub AddFiles(files As List(Of String))
        Dim sortedFiles = files.OrderBy(Function(f) f, New NaturalStringComparer()).ToList()
        For Each file In sortedFiles
            If Not inputPaths.Contains(file) Then
                inputPaths.Add(file)
            End If
        Next
        RefreshFileList()
    End Sub

    Private Sub RefreshFileList()
        lstFiles.Items.Clear()
        Dim sortedPaths = inputPaths.OrderBy(Function(p) Path.GetFileName(p), New NaturalStringComparer()).ToList()
        inputPaths.Clear()
        inputPaths.AddRange(sortedPaths)

        For Each path In inputPaths
            lstFiles.Items.Add(New FileItem(path))
        Next

        If inputPaths.Count = 0 Then
            txtFilePath.Text = "未选择任何文件"
            btnProcess.IsEnabled = False
        ElseIf inputPaths.Count = 1 Then
            txtFilePath.Text = $"已选择:{System.IO.Path.GetFileName(inputPaths(0))}"
            btnProcess.IsEnabled = True
        Else
            txtFilePath.Text = $"已选择{inputPaths.Count}个WAV文件"
            btnProcess.IsEnabled = True
        End If
        txtStatus.Text = "就绪"
    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs)
        If isProcessing Then
            MessageBox.Show("正在处理中,请等待当前任务完成", "提示", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If
        inputPaths.Clear()
        RefreshFileList()
    End Sub

    Private Sub rbDSPADPCM_Checked(sender As Object, e As RoutedEventArgs)
        selectedFormat = "--dspadpcm"
        selectedFormatName = "dspadpcm"
    End Sub

    Private Sub rbPCM16_Checked(sender As Object, e As RoutedEventArgs)
        selectedFormat = "--pcm16"
        selectedFormatName = "pcm16"
    End Sub

    Private Sub rbPCM8_Checked(sender As Object, e As RoutedEventArgs)
        selectedFormat = "--pcm8"
        selectedFormatName = "pcm8"
    End Sub

    Private Sub btnProcess_Click(sender As Object, e As RoutedEventArgs)
        If inputPaths.Count = 0 Then
            MessageBox.Show("请先选择WAV文件或文件夹", "提示", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        If Not System.IO.File.Exists(_tempExePath) Then
            MessageBox.Show("找不到NW4F_WaveConverter.exe,无法进行转换.", "错误", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        isProcessing = True
        btnProcess.IsEnabled = False
        btnSelectFiles.IsEnabled = False
        btnSelectFolder.IsEnabled = False
        btnClear.IsEnabled = False
        txtStatus.Text = $"开始处理{inputPaths.Count}个文件..."

        ProcessFilesAsync()
    End Sub

    Private Async Sub ProcessFilesAsync()
        Dim successCount As Integer = 0
        Dim failCount As Integer = 0
        Dim totalFiles = inputPaths.Count

        For i As Integer = 0 To totalFiles - 1
            Dim inputFile = inputPaths(i)
            Dim currentIndex = i + 1

            Dispatcher.Invoke(Sub()
                                  txtStatus.Text = $"正在处理({currentIndex}/{totalFiles}): {System.IO.Path.GetFileName(inputFile)}"
                              End Sub)

            Dim result = Await ConvertFileAsync(inputFile)
            If result Then
                successCount += 1
            Else
                failCount += 1
            End If
        Next

        Dispatcher.Invoke(Sub()
                              txtStatus.Text = $"处理完成:成功{successCount}个,失败{failCount}个"
                              btnProcess.IsEnabled = True
                              btnSelectFiles.IsEnabled = True
                              btnSelectFolder.IsEnabled = True
                              btnClear.IsEnabled = True
                              isProcessing = False
                          End Sub)
    End Sub

    Private Async Function ConvertFileAsync(inputFile As String) As Task(Of Boolean)
        Return Await Task.Run(Function()
                                  Try
                                      Dim directory As String = System.IO.Path.GetDirectoryName(inputFile)
                                      Dim fileNameWithoutExt As String = System.IO.Path.GetFileNameWithoutExtension(inputFile)
                                      Dim outputFileName As String = fileNameWithoutExt & "." & selectedFormatName & ".bfwav"
                                      Dim outputFilePath As String = System.IO.Path.Combine(directory, outputFileName)

                                      Dim arguments As String = $"{selectedFormat} ""{inputFile}"""

                                      Dim processInfo As New ProcessStartInfo()
                                      processInfo.FileName = _tempExePath
                                      processInfo.Arguments = arguments
                                      processInfo.UseShellExecute = False
                                      processInfo.RedirectStandardOutput = True
                                      processInfo.RedirectStandardError = True
                                      processInfo.CreateNoWindow = True
                                      processInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(_tempExePath)

                                      Using process As Process = Process.Start(processInfo)
                                          process.WaitForExit()

                                          If process.ExitCode = 0 Then
                                              Threading.Thread.Sleep(500)

                                              Dim searchPattern As String = fileNameWithoutExt & "*" & selectedFormatName & "*.bfwav"
                                              Dim possibleFiles As String() = System.IO.Directory.GetFiles(directory, searchPattern)

                                              If possibleFiles.Length > 0 Then
                                                  Dim firstFile As String = possibleFiles(0)
                                                  If firstFile <> outputFilePath Then
                                                      Try
                                                          If System.IO.File.Exists(outputFilePath) Then
                                                              System.IO.File.Delete(outputFilePath)
                                                          End If
                                                          System.IO.File.Move(firstFile, outputFilePath)
                                                      Catch
                                                      End Try
                                                  End If
                                                  Return True
                                              End If

                                              If System.IO.File.Exists(outputFilePath) Then
                                                  Return True
                                              End If
                                          End If
                                      End Using
                                  Catch ex As Exception
                                  End Try
                                  Return False
                              End Function)
    End Function

    Protected Overrides Sub OnClosed(e As EventArgs)
        Try
            If System.IO.File.Exists(_tempExePath) Then
                System.IO.File.Delete(_tempExePath)
            End If

            For Each dllPath In _tempDllPaths
                If System.IO.File.Exists(dllPath) Then
                    System.IO.File.Delete(dllPath)
                End If
            Next

            Dim tempDir = System.IO.Path.GetDirectoryName(_tempExePath)
            If System.IO.Directory.Exists(tempDir) AndAlso Not System.IO.Directory.GetFiles(tempDir).Any() Then
                System.IO.Directory.Delete(tempDir)
            End If
        Catch
        End Try

        MyBase.OnClosed(e)
    End Sub
End Class