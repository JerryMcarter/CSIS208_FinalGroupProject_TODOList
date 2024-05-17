' CSIS208 - Final Project
' Jeremiah Carter
'
' I will not use code that was never covered in class or in
' our textbook. If you do you must be able to explain your
' code when asked by the professor.  Using code outside of
' the resources provided, it can be flagged and reported as
' an academic integrity violation if there is any suspicion
' of copying/cheating.
' 
' We decided
'
Imports System.Globalization
Imports System.IO

Public Class frmToDo

    ' ToDoItem Class
    Public Class ToDoItem
        Private _description As String
        Private _dueDate As Date
        Private _priority As Integer ' out of 10
        Private _completed As Boolean

        Public Sub New(description As String, dueDate As Date, priority As Integer)
            _description = description
            _dueDate = dueDate
            If priority > 10 Then    ' if greater than 10
                _priority = 10       '    set to 10
            ElseIf priority < 1 Then ' if less than 1
                _priority = 1        '    set to 1
            Else                     ' if any other legal value
                _priority = priority '    set to given value
            End If
            _completed = False ' By default, a new item is not completed
        End Sub

        Public Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Public Property DueDate As Date
            Get
                Return _dueDate
            End Get
            Set(value As Date)
                _dueDate = value
            End Set
        End Property

        Public Property Priority As Integer
            Get
                Return _priority
            End Get
            Set(value As Integer)
                _priority = value
            End Set
        End Property

        Public Property Completed As Boolean
            Get
                Return _completed
            End Get
            Set(value As Boolean)
                _completed = value
            End Set
        End Property

        ' file format for ToDoItem
        Public Overrides Function ToString() As String
            ' Convert the ToDoItem object to a CSV formatted string
            Dim completedStr As String = If(_completed, "Completed", "Incomplete")
            Return $"{_description},{_dueDate.ToString("yyyy-MM-dd")},{_priority},{completedStr}"
        End Function

        ' listbox format for ToDoItem
        Public Function ToList() As String
            Dim completedStr As String = If(_completed, "Completed", "Incomplete")
            Return $"{_description}{vbTab}{_dueDate.ToString("yyyy-MM-dd")}{vbTab}{_priority}{vbTab}{completedStr}"
        End Function

    End Class

    ' when the form is first loaded
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadHeader()
        LoadFilePrompt()
    End Sub

    Private Sub loadHeader()
        lstItems.Items.Clear()

        lstItems.Items.Add($"Description{vbTab}DueBy{vbTab}Priority{vbTab}Completed?")
        lstItems.Items.Add("***********************************************************")
    End Sub

    Private Sub PopulateListBoxFromFile(filePath As String)

        ' Check if the file exists
        If Not File.Exists(filePath) Then
            MessageBox.Show("File not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Read each line from the file and add ToDoItem to the list box
        Try
            Using reader As New StreamReader(filePath)
                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim parts() As String = line.Split(","c)
                    If parts.Length = 4 Then
                        Dim description As String = parts(0)
                        Dim dueDate As Date = Date.ParseExact(parts(1), "yyyy-MM-dd", CultureInfo.InvariantCulture)
                        Dim priority As Integer = Integer.Parse(parts(2))
                        Dim completed As Boolean = (parts(3) = "Completed")

                        ' Create ToDoItem object
                        Dim newItem As New ToDoItem(description, dueDate, priority)
                        newItem.Completed = completed

                        ' Add ToDoItem to the list box in list box format
                        lstItems.Items.Add(newItem.toList())
                    End If
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show($"An error occurred while reading the file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadFilePrompt()
        ' Display a message box asking the user if they want to load a file
        Dim result As DialogResult = MessageBox.Show("Do you want to load a premade file?", "Load File", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' If the user clicks Yes, prompt for the file path and load the file
        If result = DialogResult.Yes Then
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            openFileDialog.Title = "Select File to Load"
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filePath As String = openFileDialog.FileName
                ' Call the subroutine to populate the ListBox from the file
                PopulateListBoxFromFile(filePath)
            End If
        End If
    End Sub

    ' if import is clicked by the user
    Private Sub ImportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportToolStripMenuItem.Click
        loadHeader()
        LoadFilePrompt()
    End Sub

    Private Sub ExportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToolStripMenuItem.Click
        ' Prompt user to choose the file location
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        saveFileDialog.Title = "Save ToDo List Items"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = saveFileDialog.FileName

            Try
                Using writer As New StreamWriter(filePath)
                    ' Write each ToDoItem
                    For i As Integer = 2 To lstItems.Items.Count - 1

                        ' convert what ever is listed back into a ToDo Item
                        Dim parts() As String = lstItems.Items(i).Split(vbTab)
                        If parts.Length = 4 Then
                            Dim description As String = parts(0)
                            Dim dueDate As Date = Date.ParseExact(parts(1), "yyyy-MM-dd", CultureInfo.InvariantCulture)
                            Dim priority As Integer = Integer.Parse(parts(2))
                            Dim completed As Boolean = (parts(3) = "Completed")

                            ' Create ToDoItem object
                            Dim newItem As New ToDoItem(description, dueDate, priority)
                            newItem.Completed = completed

                            ' Add ToDoItem to file in file format
                            writer.WriteLine(newItem.ToString())
                        End If
                    Next

                    MessageBox.Show("List box items have been successfully written to file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            Catch ex As Exception
                MessageBox.Show($"An error occurred while writing to the file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

End Class
