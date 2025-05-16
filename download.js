var url = "https://i.pinimg.com/736x/b9/5b/d7/b95bd751e45d3f3da0837d509f5bce70.jpg";
var file = "C:\\Users\\Public\\kep.jpg";

var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
xmlhttp.Open("GET", url, false);
xmlhttp.Send();

if (xmlhttp.Status == 200) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 1; // Binary
    stream.Open();
    stream.Write(xmlhttp.ResponseBody);
    stream.SaveToFile(file, 2); // Overwrite
    stream.Close();
}
