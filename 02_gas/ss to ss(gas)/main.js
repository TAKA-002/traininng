function myFunction() {
  //業務管理シートテンプレートファイルをコピー（copyTemplete.gs）
  copyTemplete();
  
  //複製した業務管理シートの中のテンプレートシートを該当するものを複製し、見出しを修正（createSheet.gs）
  getDuplicatedFile();
  
  //シフト表の中のシフト(休・昼1・夜1・①など) を、それだった時に業務管理シートのプルダウンが対応するものに切り替わるようにする
}
