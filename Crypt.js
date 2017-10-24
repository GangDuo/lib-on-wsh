var Crypt = (function() {
  function Crypt() {
    this.stream = WScript.CreateObject("ADODB.Stream");
    this.rsa = WScript.CreateObject("System.Security.Cryptography.RSACryptoServiceProvider");
  }

  /**
   * RSA アルゴリズムでデータを復号化します。
   *
   * @param text  復号化するデータ。
   * @param publicKeyAsXml RSA オブジェクトのキーを格納している XML 文字列
   * @return 暗号化する前の元のプレーン テキストである復号化されたbyte配列。
   */
  Crypt.prototype.decrypt = function(encrypted, privateKeyAsXml) {
    this.rsa.FromXmlString(privateKeyAsXml);
    try {
      this.stream.Open();
      this.stream.Type = adTypeBinary;
      this.stream.Write(encrypted);
      this.stream.Position = 0;
      return this.rsa.Decrypt(this.stream.Read(adReadAll), false);
    } catch(e) {
    } finally {
      this.stream.Close();
    }
  };

  /**
   * RSA アルゴリズムでデータを暗号化します。
   *
   * @param text  暗号化するデータ。
   * @param publicKeyAsXml RSA オブジェクトのキーを格納している XML 文字列
   * @return 暗号化されたbyte配列。
   */
  Crypt.prototype.encrypt = function(unicodeText, publicKeyAsXml) {
    this.rsa.FromXmlString(publicKeyAsXml);
    try {
      this.stream.Open();
      this.stream.Type = adTypeText;
      this.stream.Charset = Unicode.meta;
      this.stream.WriteText(unicodeText);
      this.stream.Position = 0;
      this.stream.Type = adTypeBinary;
      return this.rsa.Encrypt(this.stream.Read(adReadAll), false);
    } catch(e) {
    } finally {
      this.stream.Close();
    }
  };

  return Crypt;
})();


