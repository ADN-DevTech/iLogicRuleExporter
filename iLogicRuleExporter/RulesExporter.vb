
Imports Inventor
Imports File = System.IO.File
Imports Path = System.IO.Path

Public Class RulesExporter

    Private m_invInfo As New InventorInfo
    Private m_iLogicAuto As Object

    Private Const VersionId As String = "' ----- Internal rule exported by iLogicRuleExporter (Version 1.0)"

    Public Sub New()
        m_invInfo.StartInventor()
    End Sub


    Public Sub ExportAllRules()

        Dim ThisApplication As Application = m_invInfo.Application
        If (ThisApplication Is Nothing) Then Return
        m_iLogicAuto = GetiLogicAutomation(ThisApplication)
        Dim doc As Document = ThisApplication.ActiveDocument
        If (doc Is Nothing) Then Return

        ExportRulesInDoc(doc)
        For Each refDoc As Document In doc.AllReferencedDocuments
            ExportRulesInDoc(refDoc)
        Next

    End Sub


    Public Sub ExportRulesInDoc(ByVal doc As Document)

        Dim rules As IEnumerable = m_iLogicAuto.Rules(doc)
        If (rules Is Nothing) Then Return

        For Each rule As Object In rules
            ExportRuleFiles(doc, rule.Name, rule.Text)
        Next
    End Sub

    Sub ExportRuleFiles(doc As Document, ruleName As String, ruleText As String)
        If (String.IsNullOrEmpty(doc.FullFileName)) Then Return

        ' TODO: ReplaceIllegalCharacters might possibly map two rules to the same filename. Check for duplicates.
        Dim ruleFilePath As String = doc.FullFileName + "." + ReplaceIllegalCharacters(ruleName) + ".iLogicVb"

        ' Add an identifier at the top of the file so we can recognize it as coming from this program
        ruleText = VersionId + System.Environment.NewLine + ruleText

        If (File.Exists(ruleFilePath)) Then
            Dim oldText As String = File.ReadAllText(ruleFilePath)
            If (String.Equals(oldText, ruleText)) Then Return ' don't overwrite unless the rule has changed
            If (Not oldText.StartsWith(VersionId)) Then Return '  don't overwrite unless we created it. TODO: show warning
        End If

        File.WriteAllText(ruleFilePath, ruleText)
    End Sub

    Shared Function ReplaceIllegalCharacters(filename As String) As String
        Dim invalidChars() As Char = Path.GetInvalidFileNameChars()
        For Each chx As Char In invalidChars
            If (filename.IndexOf(chx) >= 0) Then
                filename = filename.Replace(chx, "_"c)
            End If
        Next

        Return filename
    End Function

    Friend Sub ReleaseInventor()
        m_invInfo.StopInventor()
    End Sub

End Class
