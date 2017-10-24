var Unicode = (function() {
  function Unicode() {}
  Unicode.meta = 'unicode';

  Unicode.getString = function(bytes) {
    var stream = WScript.CreateObject("ADODB.Stream");
    // binary to plain text
    try {
      stream.Open();
      stream.Type = adTypeBinary;
      stream.Write(bytes);
      stream.Position = 0;
      stream.Type = adTypeText;
      stream.Charset = Unicode.meta;
      return stream.ReadText();
    } catch(e) {
    } finally {
      stream.Close();
    }
  };

  return Unicode;
})();
