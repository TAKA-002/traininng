function myFunction() {
  //業務管理シートテンプレートファイルをコピー（copyTemplete.gs）
  copyTemplete();
  
  //複製した業務管理シートの中のテンプレートシートを複製（土日祝日と平日でことなるテンプレート（createSheet.gs）
  getDuplicatedFile();
  
  //シフト表の中のシフト(休・昼1・夜1・①など) を、それだった時に業務管理シートのプルダウンが対応するものに切り替わるようにする
}
