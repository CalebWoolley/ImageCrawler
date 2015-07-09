Imports System.Net
Imports System.IO
Imports System.Text

''' <summary>
''' An image crawling utility created by Caleb Woolley.  2015.
''' </summary>
''' <remarks>Error handling is something you should really implement on your own for this project!</remarks>
Public Class ImageCrawler
    Private homePage As String
    Private localDirectory As String
    Private cookies As New CookieContainer
    Private userAgent As String = "Mozilla/5.0 (Windows NT 6.3; rv:36.0) Gecko/20100101 Firefox/36.0"
    Private linkDatabase As New List(Of String)
    Private imageDatabase As New List(Of String)
    Private visitedLinks As New List(Of String)
    Private downloadedImages As New List(Of String)
    Private urlLimitation As String
    Private imageLimitation As String
    Private waitTime As Integer

    ''' <summary>
    ''' ImageCrawler is a class designed to scour the web for images to download.
    ''' </summary>
    ''' <param name="homeURL">The initial starting web address.</param>
    ''' <param name="downloadDirectory">Where images are to be stored.</param>
    ''' <param name="appPauseMS">How long to wait between page loads and image downloads, to prevent server overload.</param>
    ''' <param name="URLLimiter">If the links collected don't contain this limiter, they won't be crawled.</param>
    ''' <param name="imageLimiter">If the images collected don't contain this limiter, they won't be downloaded.</param>
    Public Sub New(Optional ByRef homeURL As String = "", Optional ByRef downloadDirectory As String = "", Optional ByRef appPauseMS As Integer = 100, Optional ByRef URLLimiter As String = "", Optional ByRef imageLimiter As String = "")
        homePage = homeURL
        localDirectory = downloadDirectory
        urlLimitation = URLLimiter
        imageLimitation = imageLimiter
        waitTime = appPauseMS
        If Not homePage = "" Then
            scrapeWebPage(homePage)
        End If
    End Sub

#Region "Properties"
    Public Property homeURL As String
        Get
            Return homePage
        End Get
        Set(value As String)
            scrapeWebPage(value)
            homePage = value
        End Set
    End Property

    Public Property appPauseMs As Integer
        Get
            Return waitTime
        End Get
        Set(value As Integer)
            waitTime = value
        End Set
    End Property

    Public Property URLLimiter As String
        Get
            Return urlLimitation
        End Get
        Set(value As String)
            urlLimitation = value
        End Set
    End Property

    Public Property imageLimiter As String
        Get
            Return imageLimitation
        End Get
        Set(value As String)
            imageLimitation = value
        End Set
    End Property

    Public Property downloadDirectory As String
        Get
            Return localDirectory
        End Get
        Set(value As String)
            localDirectory = value
        End Set
    End Property

    Public ReadOnly Property areLinksInDatabase As Boolean
        Get
            Return If(linkDatabase.Count = 0, False, True)
        End Get
    End Property

    Public ReadOnly Property areImagesInDatabase As Boolean
        Get
            Return If(imageDatabase.Count = 0, False, True)
        End Get
    End Property

    Public ReadOnly Property visitedLinksCount As Integer
        Get
            Return visitedLinks.Count
        End Get
    End Property

    Public ReadOnly Property downloadedImageCount As Integer
        Get
            Return downloadedImages.Count
        End Get
    End Property

    Public ReadOnly Property linkDatabaseCount As Integer
        Get
            Return linkDatabase.Count
        End Get
    End Property

    Public ReadOnly Property imageDatabaseCount As Integer
        Get
            Return imageDatabase.Count
        End Get
    End Property
#End Region

#Region "WebRequests"
    Private Function GetData(ByVal URL As String) As String
        Dim request As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
        request.CookieContainer = cookies
        request.UserAgent = userAgent
        request.Timeout = 10000
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Dim reader As New StreamReader(response.GetResponseStream())
        Return reader.ReadToEnd
    End Function

    ''' <summary>
    ''' Need to preload a website and generate some cookies?  Use this.
    ''' </summary>
    ''' <param name="URL">URL to post at.</param>
    ''' <param name="PostData">String that contains POST data.</param>
    ''' <returns>Webrequest results.</returns>
    Public Function Post(ByVal URL As String, ByVal PostData As String)
        Dim _Encryption As New UTF8Encoding
        Dim byteData As Byte() = _Encryption.GetBytes(PostData)
        Dim postReq As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
        postReq.Method = "POST"
        postReq.Timeout = 10000
        postReq.KeepAlive = True
        postReq.CookieContainer = cookies
        postReq.ContentType = "application/x-www-form-urlencoded"
        postReq.Referer = URL : postReq.UserAgent = userAgent
        postReq.ContentLength = byteData.Length : postReq.AllowAutoRedirect = True
        Dim postreqstream As Stream = postReq.GetRequestStream()
        postreqstream.Write(byteData, 0, byteData.Length)
        postreqstream.Close()
        Dim postresponse As HttpWebResponse
        postresponse = DirectCast(postReq.GetResponse(), HttpWebResponse)
        cookies.Add(postresponse.Cookies)
        Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
        Return postreqreader.ReadToEnd
    End Function
    Private Function GetImage(ByVal URL As String) As Bitmap
        Dim request As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
        request.UserAgent = userAgent
        request.CookieContainer = cookies
        request.Timeout = 10000
        Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
        Return Bitmap.FromStream(response.GetResponseStream())
    End Function
#End Region

#Region "URL Scraping"
    Private Sub scrapeWebPage(ByRef URL As String)
        On Error Resume Next
        Dim sourcePage As String = GetData(URL)
        Dim links As New List(Of String)
        links.AddRange(parseAll(sourcePage, "href=""", """"))
        links.AddRange(parseAll(sourcePage, "href='", "'"))
        For Each link As String In links
            link = If(link.StartsWith("http:"), link, MapUrl(URL, link))
            If Not linkDatabase.Contains(link) And Not visitedLinks.Contains(link) Then
                If link.Contains(urlLimitation) And link IsNot Nothing And Not link.StartsWith("javascript:") And Not link.StartsWith("#") And Not link.StartsWith("mailto:") And Not link.EndsWith(".png") And Not link.EndsWith(".jpg") And Not link.EndsWith(".gif") Then
                    linkDatabase.Add(link)
                End If
            End If
        Next

        Dim images As New List(Of String)
        images.AddRange(parseAll(sourcePage, "img src=""", """"))
        images.AddRange(parseAll(sourcePage, "img src='", "'"))
        For Each image As String In images
            image = If(image.StartsWith("http:"), image, MapUrl(URL, image))
            If image.Contains(imageLimitation) And Not imageDatabase.Contains(image) And Not downloadedImages.Contains(image) Then
                If image IsNot Nothing Then
                    imageDatabase.Add(image)
                End If
            End If
        Next
    End Sub

#Region "URL Mapping"
    Private Function MapUrl(ByVal baseAddress As String, ByVal relativePath As String) As String
        Dim u As New System.Uri(baseAddress)
        If relativePath = "./" Then
            relativePath = "/"
        End If
        If relativePath.StartsWith("/") Then
            Return u.Scheme + Uri.SchemeDelimiter + u.Authority + relativePath
        Else
            Dim pathAndQuery As String = u.AbsolutePath
            ' If the baseAddress contains a file name, like ..../Something.aspx
            ' Trim off the file name
            pathAndQuery = pathAndQuery.Split("?")(0).TrimEnd("/")
            If pathAndQuery.Split("/")(pathAndQuery.Split("/").Count - 1).Contains(".") Then
                pathAndQuery = pathAndQuery.Substring(0, pathAndQuery.LastIndexOf("/"))
            End If
            baseAddress = u.Scheme + Uri.SchemeDelimiter + u.Authority + pathAndQuery

            'If the relativePath contains ../ then
            ' adjust the baseAddress accordingly

            While relativePath.StartsWith("../")
                relativePath = relativePath.Substring(3)
                If baseAddress.LastIndexOf("/") > baseAddress.IndexOf("//" + 2) Then
                    baseAddress = baseAddress.Substring(0, baseAddress.LastIndexOf("/")).TrimEnd("/")
                End If
            End While

            Return baseAddress + "/" + relativePath
        End If
    End Function
#End Region

#Region " Parsing Functions "
    Private Function parseAll(ByRef strSource As String, ByRef strStart As String, ByRef strEnd As String, Optional ByRef startPos As Integer = 0) As List(Of String)
        Dim iPos As Integer, iEnd As Integer, strResult As String, lenStart As Integer = strStart.Length
        Dim lstAdd As New List(Of String)
        Do Until iPos = -1
            strResult = String.Empty
            iPos = strSource.IndexOf(strStart, startPos)
            iEnd = strSource.IndexOf(strEnd, iPos + lenStart)
            If iPos <> -1 AndAlso iEnd <> -1 Then
                strResult = strSource.Substring(iPos + lenStart, iEnd - (iPos + lenStart))
                lstAdd.Add(strResult)
                startPos = iPos + lenStart
            End If
        Loop
        Return lstAdd
    End Function

    Private Function parse(ByRef strSource As String, ByRef strStart As String, ByRef strEnd As String, _
                                Optional ByRef startPos As Integer = 0) As String
        Dim iPos As Integer, iEnd As Integer, lenStart As Integer = strStart.Length
        Dim strResult As String

        strResult = String.Empty
        iPos = strSource.IndexOf(strStart, startPos)
        iEnd = strSource.IndexOf(strEnd, iPos + lenStart)
        If iPos <> -1 AndAlso iEnd <> -1 Then
            strResult = strSource.Substring(iPos + lenStart, iEnd - (iPos + lenStart))
        End If
        Return strResult
    End Function

#End Region
#End Region

#Region "Crawling"
    Public Sub crawlSiteOnce()
        If homeURL = "" Then Exit Sub
        On Error Resume Next
        If linkDatabase.Count = 0 Then Exit Sub
        Threading.Thread.Sleep(waitTime)
        Dim link As String = linkDatabase(0)
        linkDatabase.RemoveAt(0)
        scrapeWebPage(link)
        visitedLinks.Add(link)
    End Sub

    Public Sub downloadLatestImage()
        If downloadDirectory = "" Then Exit Sub
        On Error Resume Next
        If imageDatabase.Count = 0 Then Exit Sub
        Threading.Thread.Sleep(waitTime)
        Dim image As String = imageDatabase(0)
        imageDatabase.RemoveAt(0)
        Dim fileName As String = image.Substring(image.LastIndexOf("/"))
        If File.Exists(localDirectory & "\" & fileName) Then
            downloadedImages.Add(image)
            Exit Sub
        End If
        Dim bm As Bitmap = GetImage(image)
        bm.Save(localDirectory & "\" & fileName)
        downloadedImages.Add(image)
    End Sub

#End Region

#Region "Continuity"
    Public Function exportLinks() As List(Of String)
        Return linkDatabase
    End Function

    Public Sub importLinks(ByRef importedList As List(Of String))
        If importedList.Count = 0 Then Exit Sub
        linkDatabase.AddRange(importedList)
    End Sub

    Public Function exportVisitedLinks() As List(Of String)
        Return visitedLinks
    End Function

    Public Sub importVisitedLinks(ByRef importedList As List(Of String))
        If importedList.Count = 0 Then Exit Sub
        visitedLinks.AddRange(importedList)
    End Sub

    Public Function exportImages() As List(Of String)
        Return imageDatabase
    End Function

    Public Sub importImages(ByRef importedList As List(Of String))
        If importedList.Count = 0 Then Exit Sub
        imageDatabase.AddRange(importedList)
    End Sub

    Public Function exportDownloadedImages() As List(Of String)
        Return downloadedImages
    End Function

    Public Sub importDownloadedImages(ByRef importedList As List(Of String))
        If importedList.Count = 0 Then Exit Sub
        downloadedImages.AddRange(importedList)
    End Sub
#End Region
End Class
