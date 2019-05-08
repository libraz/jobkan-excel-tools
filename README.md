# jobkan-excel-tools

ジョブカンの管理画面、出勤簿一括ダウンロードで取得できるexcelを対象に条件抽出を行うスクリプトです。

## インストール・設定

以下MacOS(bash系)での設定手順となります。

### Node.jsのセットアップ

`ターミナル`を開いて下記コマンドを貼り付けてください（複数行全部まとめて）

```bash
curl -L git.io/nodebrew | perl - setup

cd
cat >> .bashrc << 'EOS'
export PATH=$HOME/.nodebrew/current/bin:$PATH

bind '"\C-n": history-search-forward'
bind '"\C-p": history-search-backward'
bind '"\e[A": history-search-backward'
bind '"\e[B": history-search-forward'
EOS

cat >> .bash_profile << 'EOS'
if [ -f ~/.bashrc ] ; then
. ~/.bashrc
fi
EOS

nodebrew install-binary latest
nodebrew use latest
```

### スクリプトの取得＆配置

```bash
cd
cd Documents
mkdir jobkan-excel
cd ..
mkdir Projects
cd Projects
git clone https://github.com/libraz/jobkan-excel-tools.git
cd jobkan-excel-tools
npm install
ln -s ../../Documents/jobkan-excel ./excel
```

`git clone`実行時にMacの場合`xcode`のインストールが実行される場合があります

## excelの配置

`Finder`から`書類`以下の``jobkan-excel``にジョブカンから取得したEXCEL群を配置してください。

## スクリプトの実行

### 前準備

前準備としてスクリプトディレクトリに移動します

```bash
cd
cd Projects/jobkan-excel-tools
```

### 総労働時間判定

下記例では`労働時間（総合）`が200時間を超えているメンバーの情報を出力します

```bash
node grep.js -t 200:00
```

### 実労働時間

下記例では1日の`実労働時間`が12時間を超えているメンバーおよび勤務日の情報を出力します

```bash
node grep.js -w 12:00
```

### 退勤実時刻

下記例では`退勤実時刻`が21時を超えているメンバーおよび勤務日の情報を出力します

```bash
node grep.js -q 21:00
```

### 実残業時間

下記例では1日の`実残業時間`が3時間を超えているメンバーおよび勤務日の情報を出力します

```bash
node grep.js -o 03:00
```

### 休日出勤

休日にも関わらず`実労働時間`の記録が残っているメンバーおよび勤務日の情報を出力します

```bash
node grep.js --holiday
```

### 退勤打刻忘れ

`出勤実時刻`が入力されているにもかかわらず`退勤実時刻`が未入力のメンバーおよび勤務日の情報は常に出力されます
