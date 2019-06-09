Sub HelloWorld ()
     Dim appVisio As Visio.Application 'Instance of Visio
     Dim docsObj As Visio.Documents    'Documents collection of instance
     Dim docObj As Visio.Document      'Document to work in
     Dim stnObj As Visio.Document      'Stencil that contains master
     Dim mastObj As Visio.Master       'Master to drop
     Dim pagsObj As Visio.Pages        'Pages collection of document
     Dim pagObj As Visio.Page          'Page to work in
     Dim shpObj As Visio.Shape         'Instance of master on page

     'Create an instance of Visio and create a document based on the Basic
     'template. It doesn't matter if an instance of Visio is already running;
     'the program will run a new one.
     Set appVisio = CreateObject("visio.application")
     Set docsObj = appVisio.Documents
     'Create a document based on the Basic Diagram template which automatically
     'opens the Basic Shapes stencil.
     Set docObj = docsObj.Add("基本ダイアグラム.vst")
     Set pagsObj = appVisio.ActiveDocument.Pages
     'A new document always has at least one page, whose index in the Pages collection is 1.
     Set pagObj = pagsObj.Item(1)
     Set stnObj = appVisio.Documents("基本シェイプ.vss")
     Set mast0bj = stn0bj.Masters("長方形")
     'Drop the rectangle in the approximate middle of the page.
     'Coordinates passed with Drop are always inches.
     Set shpObj = pagObj.Drop(mastObj, 4.25, 5.5)
     'Set the text of the rectangle
     shpObj.Text = "Hello World!"
     'Save the drawing and quit Visio. The message pauses the program
     'so you can see the Visio drawing before the instance closes.
     docObj.SaveAs "hello.vsd"
     MsgBox "Drawing finished!", , "Hello World!"
     appVisio.Quit
End Sub
