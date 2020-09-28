Imports System
Imports System.Windows.Forms

Public Class WaitCursor
    Implements IDisposable

    ReadOnly savedCursor As Cursor

    Public Sub New()
        savedCursor = Cursor.Current
        Cursor.Current = Cursors.WaitCursor
    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                Cursor.Current = savedCursor
            End If
        End If
        disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
    End Sub
#End Region

End Class
