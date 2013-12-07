' GetGenre.vbs
' Author: Jonathan Aherne
' Call last.fm to populate genre field
' Written in 2012, updated in 2013
' Version 1.2

Sub GetGenre
  ' Vars
  Dim list, itm, i, genre, artist, api
  
  ' My API Key
  Set api = "aaaaaaaaaaaaaaaaaaaaaaaa"
  
  ' Get list of selected tracks
  Set list = SDB.CurrentSongList

  ' Process tracks
  For i=0 to list.count-1
    Set itm = list.Item(i)
	
    ' Grab genre from internet
    genre = itm.Genre
    artist = itm.ArtistName
  
    Dim doc, xml, nodes, node, plot
    Set doc = CreateObject("MSXML2.XMLHTTP")
    Dim baseUrl, artistUrl, api, fullUrl
  
    ' Create Url
    baseUrl = "http://ws.audioscrobbler.com/2.0/?method=artist.gettoptags"
    artistUrl = "&artist=" & artist
    api = "&api_key=" & api
    fullUrl = baseURL & artistUrl & api
  
    ' Get XML document
    doc.open "GET", fullUrl, False
    doc.send
    Set xml = CreateObject("MSXML2.DOMDocument")
    xml.load(doc.responsestream)
  
    ' Set new genre
    Set nodes = xml.getElementsByTagName("name")
    Set node = nodes.item(0)
    node.Text = ucase(left(node.Text,1)) + right(node.Text, len(node.Text)-1)
    itm.Genre = node.Text
	  
  Next
  
  ' Write and save
  list.UpdateAll
End Sub
