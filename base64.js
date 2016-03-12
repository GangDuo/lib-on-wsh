var base64 = (function() {
  var adTypeBinary = 1;
  var adTypeText = 2;
  var UTF8 = 'UTF-8';

  return {
    encode: _encode,
    decode: _decode
  };

  /**
   * 文字列をbase64でエンコードする
   *
   * @param text  base64でエンコードする文字列
   * @param arguments[1] charset string（省略可能）
   *                     "UTF-8", "Shift-JIS"
   *                     未指定の場合は"UTF-8"
   * @return base64でエンコードされた文字列
   */
  function _encode(text) {
    var enc = arguments[1] || UTF8,
        base64Text;

    _usingStream(function(stream) {
      // write
      stream.Charset = enc;
      stream.Type = adTypeText;
      stream.WriteText(text);
      // read
      stream.Position = 0;
      stream.Type = adTypeBinary;
      if(enc === UTF8) {
        stream.Position = 3; // BOM スキップ
      }
      base64Text = stringify(stream.Read(), 'bin.base64').replace(/\n/g, '');
    });
    return base64Text;
  }

  /**
   * base64でエンコードされた文字列をデコードする
   *
   * @param base64Text base64でエンコードされた文字列
   * @param arguments[1] charset string（省略可能）
   *                     "UTF-8", "Shift-JIS"
   *                     未指定の場合は"UTF-8"
   * @return デコードされた文字列
   */
  function _decode(base64Text) {
    var enc = arguments[1] || UTF8,
        text;

    _usingStream(function(stream) {
      // write
      stream.Charset = enc;
      stream.Type = adTypeBinary;
      stream.Write(parse(base64Text, 'bin.base64'));
      // read
      stream.Position = 0;
      stream.Type = adTypeText;
      text = stream.ReadText();
    });
    return text;
  }

  /**
   * ADODB.Streamのテンプレート
   * OpenしてCloseする
   *
   * @param proc ADODB.Stream -> any
   * @return なし
   */
  function _usingStream(proc) {
    var stream;

    try {
      stream = WScript.CreateObject('ADODB.Stream');
      stream.Open();
      proc.call(this, stream);
    } catch(e) {
      throw new Error(e.message);
    } finally {
      try {
        stream.Close();
      } catch(e) {}
      stream = null;
    }
  }

  /**
   * 「byte配列」から「文字列」に変換
   *
   * @param bytes
   * @param type 変換形式
   *             string、number、Int、Boolean、dateTime、bin.hex、bin.base64
   * @return 文字列
   */
  function stringify(bytes, type) {
    var text;

    _usingDomDocument(type, function(element) {
      element.nodeTypedValue = bytes;
      text = element.text;
    });
    return text;
  }

  /**
   * 「文字列」から「byte配列」に変換
   *
   * @param text
   * @param type 変換形式
   *             string、number、Int、Boolean、dateTime、bin.hex、bin.base64
   * @return byte配列
   */
  function parse(text, type) {
    var bytes;

    _usingDomDocument(type, function(element) {
      element.text = text;
      bytes = element.nodeTypedValue;
    });
    return bytes;
  }

  /**
   */
  function _usingDomDocument(type, proc) {
    var doc,
        element;

    try {
      doc = WScript.CreateObject('MSXML2.DOMDocument.6.0');
      element = doc.createElement('element');
      element.dataType = type;
      proc.call(this, element);
    } catch(e) {
      throw new Error(e.message);
    } finally {
      element = null;
      doc = null;
    }
  }
})();
