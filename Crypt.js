var Crypt = (function() {
  function Crypt() {
    this.stream = WScript.CreateObject("ADODB.Stream");
    this.rsa = WScript.CreateObject("System.Security.Cryptography.RSACryptoServiceProvider");
  }

  /**
   * RSA �A���S���Y���Ńf�[�^�𕜍������܂��B
   *
   * @param text  ����������f�[�^�B
   * @param publicKeyAsXml RSA �I�u�W�F�N�g�̃L�[���i�[���Ă��� XML ������
   * @return �Í�������O�̌��̃v���[�� �e�L�X�g�ł��镜�������ꂽbyte�z��B
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
   * RSA �A���S���Y���Ńf�[�^���Í������܂��B
   *
   * @param text  �Í�������f�[�^�B
   * @param publicKeyAsXml RSA �I�u�W�F�N�g�̃L�[���i�[���Ă��� XML ������
   * @return �Í������ꂽbyte�z��B
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


