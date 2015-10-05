' Program: Business Rule & Data Classes
' Author:  Robert Peck
' Date:    03/03/2015

Option Explicit On
Option Strict On
Imports DataBaseTables.Tables

Public Class DataBaseForm

    Dim TheConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=J:\CIS212\RobertPeckProject02\WindowsApplication1\Northwind.mdb;"
    Private aDataRun As New DataBaseTables.Tables.DataBaseTableSelection(TheConnectionString)
    Private UnitPrice As Decimal
    Private IDBoolean As Boolean = False
    Private IDInput As Integer = Nothing

#Region "   Listbox/Grid Filled via filter textbox   "

    Private Sub FilterButton_Click(sender As Object, e As EventArgs) Handles FilterButton.Click 'When Filter Button Is Clicked Program will Pull Information from the database based on the input filter


        Try
            DataGridView1.DataSource = (aDataRun.FillGrid(FilterTextBox.Text)).Tables(0)
            ListBox1.Items.Clear()
            With aDataRun.RetrieveProductNamePrice(FilterTextBox.Text)
                If .HasRows Then
                    Do While .Read
                        ListBox1.Items.Add(.GetString(0))
                    Loop
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(aDataRun.ErrorMessage)
        End Try
        '    If (aDataRun.GetData(FilterTextBox.Text)) Then  ' Sends the filter over to the business class
        '        DataGridView1.DataSource = aDataRun.DataSet.Tables(0)       ' Sets the datasource to fill the datebase into the datagrid
        '        Try
        '            ListBox1.Items.Clear()
        '            With aDataRun.DataReader    ' Pulls the data reader to read date for writing into the list box
        '                If .HasRows Then    'If data reader still has rows it will continue
        '                    Do While .Read  'Loops through and reads the information for writing
        '                        ListBox1.Items.Add(.GetString(0))   'adds the item to the list box
        '                    Loop    ' ends loop
        '                Else
        '                    MessageBox.Show("No Data")  ' Message for when there is no data to write
        '                End If
        '            End With
        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)     ' Error message if there is an error while running the data reader
        '        End Try
        '    Else
        '        MessageBox.Show(aDataRun.ErrorMessage)  ' Error message if there is an error while writing to the data grid
        '    End If
    End Sub
#End Region

#Region " Fill ListBox w/ DataSet "
    Private Sub ListBox1_DoubleClick(sender As Object, e As System.EventArgs) Handles ListBox1.DoubleClick
        ' Pulls the Product ID number when double clicking an item in the list box
        If ListBox1.SelectedValue Is Nothing Then
            MessageBox.Show("Selected value is NULL")   ' If there is no Product ID the message will display null
        Else
            MessageBox.Show(ListBox1.SelectedValue.ToString)    ' Sets the Product ID into a message box
        End If
    End Sub
    ' When Bind Button is clicked the program will pull all the products in the database into the list box
    Private Sub BindButton_Click(sender As Object, e As EventArgs) Handles BindButton.Click
        Try


            Select Case True
                Case aDataRun.FillListBox().Tables().Count = 0
                    MessageBox.Show("Nothing Found. Table count: 0")
                Case aDataRun.FillListBox().Tables(0).Rows.Count = 0
                    MessageBox.Show("Nothing Found. Row Count: 0")
                Case Else
                    ListBox1.DisplayMember = "ProductName"
                    ListBox1.ValueMember = "ProductID"
                    ListBox1.DataSource = aDataRun.FillListBox().Tables(0)
            End Select
        Catch ex As Exception
            MessageBox.Show(aDataRun.ErrorMessage)
        End Try


        'If (aDataRun.BindListBox()) Then    ' runs the method to pull information from the database
        '    Select Case True
        '        Case aDataRun.DataSet.Tables().Count = 0
        '            MessageBox.Show("Nothing Found. Table count: 0") ' if the data set count is 0 the following message will display
        '        Case aDataRun.DataSet.Tables(0).Rows.Count = 0
        '            MessageBox.Show("Nothing Found. Row Count: 0") ' if the row count is 0 the following message will display
        '        Case Else
        '            ListBox1.DisplayMember = "ProductName"  ' By default if the table and rows are not 0 the database data will be put into the list box
        '            ListBox1.ValueMember = "ProductID"
        '            ListBox1.DataSource = aDataRun.DataSet.Tables(0)
        '    End Select
        'Else
        '    MessageBox.Show(aDataRun.ErrorMessage)  ' If an error occurs in filling the list box the following message will be display
        'End If
    End Sub
#End Region

#Region " Insert Data Into Database "
    ' When the insert button is clicked the program will attempt to add the data into the database
    Private Sub InsertButton_Click(sender As Object, e As EventArgs) Handles InsertButton.Click
        If ValidateInput() Then ' Runs a input validation method to make sure the inputs are valid
            If (aDataRun.CreateData(ProductNameTextBox.Text, UnitPrice)) = 1 Then    ' runs method to insert the data and sends over the product name and price
                MessageBox.Show("Data Has Been Added Successfully", "Success", MessageBoxButtons.OK)    ' if insert was successfull the following message will display
            Else
                MessageBox.Show(aDataRun.ErrorMessage)  ' if insert faile the following message will display
            End If
        End If
    End Sub

#End Region

#Region " Update Data In Database "
    ' Update button will update an existing record within the data base, the id will be required to update the record
    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        IDBoolean = True    ' Boolean to add the additional validation for validating the id
        If ValidateInput() Then ' runs the validate input method
            If (aDataRun.UpdateData(ProductNameTextBox.Text, UnitPrice, IDInput)) = 1 Then   ' if validation successful the program will send over the unitprice and id to update the information
                MessageBox.Show("Data Has Been Updated Successfully", "Success", MessageBoxButtons.OK)  ' If update is successful the following message will display
            Else
                MessageBox.Show(aDataRun.ErrorMessage)  ' If update fails the following message will display
            End If
        End If
        IDBoolean = False   ' resets the id boolean for validation back to false
    End Sub

#End Region

#Region " Delete Data From Database "
    ' Delete button will delete data from the database, the id will be required to delete a record
    Private Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles DeleteButton.Click
        IDBoolean = True    ' Boolean to add the additional validation for validating the id
        If ValidateInput() Then ' runs the validate input method
            If (aDataRun.DeleteData(ProductNameTextBox.Text, IDInput)) = 1 Then ' if validation successful the program will send over the Product name and id to delete the information
                MessageBox.Show("Data Was Delted Successfully", "Success", MessageBoxButtons.OK)    ' If the delete is successful the following message will display
            Else
                MessageBox.Show(aDataRun.ErrorMessage)  ' If Delete fails the following message will display
            End If
        End If
        IDBoolean = False   ' resets the id boolean for validation back to false
    End Sub

#End Region

#Region " Input Validation "
    ' Method for validating the user input
    Private Function ValidateInput() As Boolean
        If ProductNameTextBox.Text.Trim <> String.Empty Then    ' Makes sure the product name text box is filled
            Try
                UnitPrice = Decimal.Parse(UnitPriceTextBox.Text)    ' Parses the unit price and stores it in a variable
                If IDBoolean Then
                    Try
                        IDInput = Integer.Parse(IDTextBox.Text) ' parses the ID and stores it in a variable
                        Return True ' Returns true to show data is validated
                    Catch ex As Exception
                        MessageBox.Show("Please Enter A Vaild ID Number For Product", "ID Number Error", MessageBoxButtons.OK)
                        With IDTextBox      ' Displays message to user that this input needs to be fixed
                            .SelectAll()
                            .Focus()
                        End With
                        Return False    ' Returns false if the data input is incorrect
                    End Try
                End If
                Return True ' Returns true when ID isn't in use and data is validated
            Catch ex As Exception
                MessageBox.Show("Please Enter A Unit Price As A Decimal", "Unit Price Error", MessageBoxButtons.OK)
                With UnitPriceTextBox   ' Displays message to user that this input needs to be fixed
                    .SelectAll()
                    .Focus()
                End With
                Return False    ' Returns false if the data input is incorrect
            End Try
        Else
            MessageBox.Show("Please Enter A Product Name", "Product Name Error", MessageBoxButtons.OK)
            With ProductNameTextBox ' Displays message to user that this input needs to be fixed
                .SelectAll()
                .Focus()
            End With
            Return False    ' Returns false if the data input is incorrect
        End If
    End Function

#End Region

    Private Sub ClearDataButton_Click(sender As Object, e As EventArgs) Handles ClearDataButton.Click   ' Resets the textboxes to empty
        ProductNameTextBox.Text = String.Empty
        UnitPriceTextBox.Text = String.Empty
        IDTextBox.Text = String.Empty
        FilterTextBox.Text = String.Empty
        ProductNameTextBox.Focus()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Close() ' Closes the program
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
