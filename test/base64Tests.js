WScript.Echo('QWxhZGRpbjpvcGVuIHNlc2FtZQ=='===base64.encode('Aladdin:open sesame'));
WScript.Echo('Aladdin:open sesame' === base64.decode('QWxhZGRpbjpvcGVuIHNlc2FtZQ=='));
WScript.Echo('44GC' === base64.encode('あ'));
WScript.Echo('gqA=' === base64.encode('あ', 'Shift-JIS'));
WScript.Echo('あ' === base64.decode('44GC'));
WScript.Echo('あ' === base64.decode('gqA=', 'Shift-JIS'));
