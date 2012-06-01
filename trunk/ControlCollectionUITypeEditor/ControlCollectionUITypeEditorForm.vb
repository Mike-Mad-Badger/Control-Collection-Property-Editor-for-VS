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

Public Class ControlCollectionUITypeEditorForm
    'Used to allow the property editor to test if cancel was clicked
    'Could probably be done using DialogResult also
    Public IsCancelled As Boolean = False

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.IsCancelled = False
        'Hide the form, don't close it, so the property editor can refer to the selections made
        Me.Hide()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.IsCancelled = True
        'Hide the form, don't close it, so the property editor can refer to the selections made
        Me.Hide()
    End Sub

End Class