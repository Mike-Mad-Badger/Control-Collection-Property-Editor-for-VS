'This file is part of ControlCollectionUITypeEditor.

'Copyright 2012 Mike Conroy

'ControlCollectionUITypeEditor is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'ControlCollectionUITypeEditor is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with Foobar.  If not, see <http://www.gnu.org/licenses/>.

Imports System.Collections.ObjectModel
Imports System.Windows.Forms

<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>
Public Class ControlCollectionUITypeEditor
    Inherits AbstractUITypeEditor

    'This is the control to be used in design time DropDown editor
    Private WithEvents myControl As ControlCollectionUITypeEditorForm

    'See base class for comments about this method
    Protected Overrides Function GetEditControl(ByVal PropertyName As String, ByVal CurrentValue As Object) As Control
        'create the control
        myControl = New ControlCollectionUITypeEditorForm()
        'it's a form in this case so set its title
        myControl.Text = "Edit Control Array Property: " & PropertyName
        'give a reference to the control back to the base class method
        'the control has not been shown at this point, just created
        Return myControl
    End Function

    'See base class for comments about this method
    Protected Overrides Function GetEditedValue(ByVal EditControl As Control, ByVal PropertyName As String, ByVal OldValue As Object) As Object
        'if for some reason no control has been created just return the previous (current) value for the property
        If EditControl Is Nothing Then Return OldValue
        'Check that the cancel button hasn't been pressed on the control
        If myControl.IsCancelled Then
            Return OldValue
        Else
            'Return the new value for the property

            'first test if anything has been checked
            If myControl.AvailableControls.CheckedItems.Count = 0 Then
                'if nothing has been selected in the CheckedListBox then return null
                Return Nothing
            Else
                'otherwise copy the Controls represented by the checked items in the CheckedListBox to a Collection
                'the Collection is the same type as the property in the custom component

                'create a temporary Collection
                Dim tCollection As New Collection(Of Control)
                For Each c As Control In myControl.AvailableControls.CheckedItems
                    'cycle through the checked controls only and add each one to the temporary collection
                    tCollection.Add(c)
                Next
                'return the temporary collection to the property
                Return tCollection
            End If
        End If
    End Function

    Protected Overrides Sub LoadValues(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As System.IServiceProvider, ByVal value As Object)
        'Load the AvailableControls checked list with all form controls visible through the Designer
        For Each c As Component In context.Container.Components
            'Cycle through the components owned by the form in the designer
            If TypeOf c Is Control Then
                'if the component is a control then add it to the CheckedListBox
                myControl.AvailableControls.Items.Add(c)
            End If
        Next
        If value IsNot Nothing Then
            'If the property currently has anything in its collection then check those items in the checked list box
            'create a temporary Collection that holds the current controls held by the property
            Dim tCollection As Collection(Of Control) = CType(value, Collection(Of Control))
            Dim found As Integer
            For Each c As Control In tCollection
                'cycle through the current controls held by the property
                found = myControl.AvailableControls.Items.IndexOf(c)
                If Not found = -1 Then
                    'if the control has been found in the CheckedListBox then set its cehcked state to true
                    myControl.AvailableControls.SetItemChecked(found, True)
                End If
            Next
        End If
    End Sub

End Class
