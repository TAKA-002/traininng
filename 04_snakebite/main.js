"use strict";

let W = 20;
let H = 20;
let S = 20;

let snake = [];
let foods = [];

let keyCode = 0;
let point = 0;
let timer = NaN;
let ctx;

//Pointオブジェクトを定義しておく。
//現在値の値を取得しているオブジェクト。
function Point(x, y) {
  this.x = x;
  this.y = y;
}

//=================================
//初期化
//=================================
function init() {
  //bodyのcanvasタグを取得
  let canvas = document.getElementById("field");

  //canvasのwidth400/20の20をWに格納
  //canvas領域は20×20になる
  W = canvas.width / S;
  H = canvas.height / S;

  //ctxにcanvasのコンテキストオブジェクトを生成。
  //描画するために用意する道具は2dを書くためのものやでって言っている。
  ctx = canvas.getContext("2d");

  //ctxの２dのものを使って、書くフォントはこれっていう指定
  ctx.font = "20px sans-serif";

  //snakeの初期化
  //先に定義していたポイントオブジェクトをインスタンス化して使用。
  //仮引数xに20/2の10　仮引数yにも同じく渡す　▶　中心座標
  //それを配列に追加
  snake.push(new Point(W / 2, H / 2));

  //餌の追加
  //初期値iを0として、最大10回addFoodメソッドを繰り返し実行　▶　10個の餌が初期値
  for (let i = 0; i < 10; i++) {
    addFood();
  }

  //200ミリ秒ごとにtickメソッドを実行し、
  timer = setInterval("tick()", 200);
  window.onkeydown = keydown;
}

//=================================
//餌の追加
//=================================
function addFood() {
  while (true) {
    //xとyの値を乱数で設定し、その場所に餌を追加するということ。
    let x = Math.floor(Math.random() * W); //randam値×20(ここのWはおそらくグローバル)で切り捨て。
    let y = Math.floor(Math.random() * H);

    //ただしそれが餌が今もあるところか、または蛇がいるところだとまずい。
    //だから、isHItメソッドを作成してどちらかがtrueだったら、continueで、乱数設定からやりなおし。
    if (isHit(foods, x, y) || isHit(snake, x, y)) {
      continue;
    }

    //falseなら（Hitしなければ）餌を追加。
    //餌を追加したらbreakで終了
    //終了したら呼び出し元のfor文へ戻る。(10回実施する。）
    foods.push(new Point(x, y));
    break;
  }
}

//=================================
//衝突判定
//=================================
function isHit(data, x, y) {
  //dataはfoodかsnakeの配列の中に、Pointオブジェクトがあるか確認している。
  //snakeかfoodがdataに入り、その配列の数の回数for文で繰り返し処理。
  //実施する処理は、dataの０番目のxがxであるか。yがyであるか。
  //そうならtrueを返す。
  //違ったらfalseを返す。
  //つまり存在確認。

  for (let i = 0; i < data.length; i++) {
    if (data[i].x == x && data[i].y == y) {
      return true;
    }
  }
  return false;
}

//=================================
//蛇が餌とヒットしたら餌を動かす。
//=================================
function moveFood(x, y) {
  foods.filter((p) => {
    return p.x != x || p.y != y;
  });
}

//=================================
//メインループ
//=================================
function tick() {
  let x = snake[0].x;
  let y = snake[0].y;

  switch (keyCode) {
    case 37:
      x--; //左
      break;
    case 38:
      y--; //上
      break;
    case 39:
      x++; //右
      break;
    case 40:
      y++; //下
      break;
    default:
      paint();
      return;
  }

  //自分もしくは壁に衝突したら。
  if (isHit(snake, x, y) || x < 0 || x >= W || y < 0 || y >= H) {
    clearInterval(timer);
    paint();
    return;
  }

  //頭を戦闘に追加
  snake.unshift(new Point(x, y));

  if (isHit(foods, x, y)) {
    point += 10; //餌をたべた
    moveFood(x, y);
  } else {
    snake.pop();
  }

  paint();
}

//=================================
//ペイント
//=================================
function paint() {
  ctx.clearRect(0, 0, W * S, H * S);
  ctx.fillStyle = "rgb(256,0,0)";
  ctx.fillText(point, S, S * 2);
  ctx.fillStyle = "rgb(0,0,255)";

  foods.forEach((p) => {
    ctx.fillText("+", p.x * S, (p.y + 1) * S);
  });

  snake.forEach((p) => {
    ctx.fillText("*", p.x * S, (p.y + 1) * S);
  });
}

function keydown(e) {
  keyCode = e.keyCode;
}
