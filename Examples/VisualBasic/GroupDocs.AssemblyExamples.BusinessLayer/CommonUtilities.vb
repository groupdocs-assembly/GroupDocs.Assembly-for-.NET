
Imports GroupDocs.Assembly
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace GroupDocs.AssemblyExamples.BusinessLayer
    'ExStart:CommonUtilities
    Public Class CommonUtilities

        Public Const sourceFolderPath As String = "../../../../Data/Source/"
        Public Const destinationFolderPath As String = "../../../../Data/Destination/"
        Public Const licensePath As String = "../../GroupDocs.Assembly Product Family.lic"

#Region "DocumentDirectories"

        'ExStart:DocumentDirectories
        ''' <summary>
        ''' Takes source file name as argument. 
        ''' </summary>
        ''' <param name="sourceFileName">Source file name</param>
        ''' <returns>Returns explicit path by combining source folder path and source file name.</returns>

        Public Shared Function GetSourceDocument(sourceFileName As String) As String
            Return Path.Combine(Path.GetFullPath(sourceFolderPath), sourceFileName)
        End Function
        ''' <summary>
        ''' Takes output file name as argument. 
        ''' </summary>
        ''' <param name="outputFileName">output file name</param>
        ''' <returns>Returns explicit path by combining destination folder path and output file name.</returns>
        Public Shared Function SetDestinationDocument(outputFileName As String) As String
            Return Path.Combine(Path.GetFullPath(destinationFolderPath), outputFileName)
        End Function
        'ExEnd:DocumentDirectories
#End Region

#Region "ProductLicense"

        ''' <summary>
        ''' Set product's license
        ''' </summary>

        Public Shared Sub ApplyLicense()
            Dim lic As New License()
            lic.SetLicense(licensePath)
        End Sub

#End Region
    End Class
    'ExEnd:CommonUtilities
End Namespace


