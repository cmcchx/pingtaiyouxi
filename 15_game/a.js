var fso = new ActiveXObject("Scripting.FileSystemObject");
var a = fso.CreateTextFile("d:\\testfile.txt", true);
a.WriteLine("This is a test.");
a.Close();