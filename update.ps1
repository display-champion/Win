# =============================================================
#  ニンテンドーストア「オンラインで協力」一覧 更新スクリプト
#  - Nintendo公式ソフト検索APIから対象タイトルを取得
#  - Deku Deals(日本eShop)から過去最安値を取得
#  - index.html を再生成
#  - (任意) GitHub Pages(Winリポジトリ)へ push して公開
#
#  使い方(通常はbatから起動):
#    powershell -ExecutionPolicy Bypass -File update.ps1            # フル更新
#    powershell -ExecutionPolicy Bypass -File update.ps1 -SkipDeku  # 価格のみ高速更新(過去最安は前回値を流用)
#    powershell -ExecutionPolicy Bypass -File update.ps1 -NoPush    # 公開せずローカル生成のみ
# =============================================================
param(
    [switch]$SkipDeku,   # 過去最安の再取得をスキップ(前回キャッシュを流用)
    [switch]$NoPush,     # GitHub Pagesへの公開をしない
    [switch]$Yes         # 公開確認プロンプトを自動でYes
)

$ErrorActionPreference = 'Stop'
$root   = $PSScriptRoot
$dataDir= Join-Path $root 'data'
$UA     = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36'
$today  = (Get-Date).ToString('yyyy-MM-dd')
if(-not (Test-Path $dataDir)){ New-Item -ItemType Directory -Path $dataDir | Out-Null }

function Log($m){ Write-Host "  $m" }
Write-Host "===================================================="
Write-Host " オンラインで協力 一覧 更新  ($today)"
Write-Host "===================================================="

# ---------- 1) Nintendo公式APIからSwitchソフト全件→ONLINE_COOP抽出 ----------
Write-Host "`n[1/4] Nintendo公式APIから対象タイトルを取得中..."
$base   = 'https://search.nintendo.jp/nintendo_soft/search.json'
$limit  = 400
$matched= New-Object System.Collections.ArrayList
$page   = 1
$total  = 999999
while(($page-1)*$limit -lt $total){
    $u = "$base`?limit=$limit&page=$page&opt_hard=1_HAC&sort=score"
    try   { $r = Invoke-RestMethod -Uri $u -Headers @{ 'User-Agent'=$UA } }
    catch { Start-Sleep -Milliseconds 700; $r = Invoke-RestMethod -Uri $u -Headers @{ 'User-Agent'=$UA } }
    $total = $r.result.total
    foreach($it in $r.result.items){
        if("$($it.sctg)" -eq 'aoc' -or "$($it.sform)" -eq 'DLC'){ continue }   # DLC(追加コンテンツ)は除外
        [void]$matched.Add($it)
    }
    Write-Progress -Activity "Nintendo API" -Status "page $page (collected=$($matched.Count))" -PercentComplete ([Math]::Min(100, ($page*$limit/$total*100)))
    $page++
}
Write-Progress -Activity "Nintendo API" -Completed
# dedupe by id, build records
$seen=@{}; $recs=New-Object System.Collections.ArrayList
foreach($it in $matched){
    $key="$($it.id)"
    if($seen.ContainsKey($key)){ continue }
    $seen[$key]=$true
    $genre = if($it.genre){ ($it.genre -join ' / ') } else { '' }
    $np    = if($it.nplayer){ ($it.nplayer -join ', ') } else { '' }
    [void]$recs.Add([pscustomobject]@{
        title=[string]$it.title; maker=[string]$it.maker; genre=$genre; nplayer=$np
        price=$it.price; cur=$it.current_price; dprice=$it.dprice; pprice=$it.pprice
        sale=[int]$it.sale_flg; ssitu=[string]$it.ssitu
        img="https://img-eshop.cdn.nintendo.net/i/$($it.iurl).jpg"
        url="https://store-jp.nintendo.com/list/software/$($it.nsuid).html"
        desc=[string]$it.hcopy; nsuid=[string]$it.nsuid; low=$null
        coop=$(if($it.tags -and ($it.tags -match 'ONLINE_COOP')){ 1 } else { 0 })
    })
}
Log "対象タイトル: $($recs.Count) 件 (うちオンラインで協力: $((@($recs | Where-Object { $_.coop -eq 1 })).Count) 件)"

# ---------- 2) Deku Dealsから過去最安値(日本eShop)を取得 ----------
$lowCache = Join-Path $dataDir 'deku_lows.json'
$lowMap = @{}
if($SkipDeku){
    Write-Host "`n[2/4] 過去最安の取得はスキップ(前回キャッシュを流用)"
    if(Test-Path $lowCache){
        $c = Get-Content $lowCache -Raw -Encoding utf8 | ConvertFrom-Json
        foreach($p in $c.PSObject.Properties){ if($p.Value.low){ $lowMap[[string]$p.Value.nsuid]=[int]$p.Value.low } }
        Log "キャッシュ読込: $($lowMap.Count) 件"
    } else { Log "キャッシュが無いため過去最安は空になります" }
}
else{
    Write-Host "`n[2/4] Deku Dealsから過去最安値(日本eShop)を取得中... (数分〜十数分かかります)"
    $sess=New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $sess.Cookies.Add((New-Object System.Net.Cookie('country','JP','/','.dekudeals.com')))
    function CleanTitle($t){ ($t -replace "[™®℠]","" -replace "\s*[:：\-–—|/].*$","" -replace "\s*[\(（\[【].*$","").Trim() }
    function Lookup($title,$nsuid){
        foreach($q in (@($title,(CleanTitle $title)) | Select-Object -Unique)){
            if(-not $q){ continue }
            $enc=[uri]::EscapeDataString($q)
            try{ $r=Invoke-WebRequest -Uri "https://www.dekudeals.com/search?q=$enc" -UserAgent $UA -WebSession $sess -UseBasicParsing -TimeoutSec 25 }catch{ Start-Sleep -Milliseconds 800; continue }
            $slugs=[regex]::Matches($r.Content,"/items/[a-z0-9][a-z0-9-]+") | ForEach-Object { $_.Value } | Select-Object -Unique | Select-Object -First 5
            foreach($s in $slugs){
                try{ $p=Invoke-WebRequest -Uri "https://www.dekudeals.com$s" -UserAgent $UA -WebSession $sess -UseBasicParsing -TimeoutSec 25 }catch{ Start-Sleep -Milliseconds 400; continue }
                if($p.Content -match [regex]::Escape($nsuid)){
                    $m=[regex]::Match($p.Content,"<script id='price_history_data'[^>]*>(.*?)</script>","Singleline")
                    if($m.Success){
                        try{ $hist=($m.Groups[1].Value | ConvertFrom-Json).data }catch{ $hist=$null }
                        $prices=@($hist | ForEach-Object { $_[2] } | Where-Object { $_ -ne $null })
                        if($prices.Count){ return @{status='ok';low=[int]($prices|Measure-Object -Minimum).Minimum} }
                    }
                    return @{status='match_nohist'}
                }
                Start-Sleep -Milliseconds 200
            }
        }
        return @{status='no_match'}
    }
    # 過去最安の照会は「オンラインで協力」対応タイトルのみ(全件照会は数時間かかるため)
    $coopRecs=@($recs | Where-Object { $_.coop -eq 1 })
    Log "照会対象(オンラインで協力のみ): $($coopRecs.Count) 件"
    $store=@{}; $i=0; $ok=0
    foreach($g in $coopRecs){
        $i++
        $res=Lookup $g.title $g.nsuid
        $store[[string]$g.nsuid]=@{ title=$g.title; nsuid=$g.nsuid; status=$res.status; low=$res.low }
        if($res.status -eq 'ok'){ $ok++; $lowMap[[string]$g.nsuid]=[int]$res.low }
        if($i % 10 -eq 0 -or $i -eq $coopRecs.Count){
            Write-Progress -Activity "Deku Deals" -Status "$i/$($coopRecs.Count)  取得済 $ok" -PercentComplete ($i/$coopRecs.Count*100)
            $store | ConvertTo-Json -Depth 4 | Out-File $lowCache -Encoding utf8
        }
        Start-Sleep -Milliseconds 250
    }
    Write-Progress -Activity "Deku Deals" -Completed
    $store | ConvertTo-Json -Depth 4 | Out-File $lowCache -Encoding utf8
    Log "過去最安を取得: $ok / $($coopRecs.Count) 件"
}
# merge low into recs
$mergeCnt=0
foreach($g in $recs){ if($lowMap.ContainsKey([string]$g.nsuid)){ $g.low=$lowMap[[string]$g.nsuid]; $mergeCnt++ } }
Log "過去最安を付与: $mergeCnt 件"

# ---------- 3) index.html を再生成 ----------
Write-Host "`n[3/4] index.html を生成中..."
$recsJson = ($recs | ConvertTo-Json -Depth 4 -Compress)
$tpl = @'
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nintendo Switch ゲーム一覧（オンライン協力フィルタ付き）</title>
<style>
  :root{ --bg:#f2f4f7; --card:#fff; --ink:#1b2733; --sub:#5b6b7b; --line:#e3e8ef; --red:#e60012; --accent:#0a84ff; --low:#0a9d58; }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:"Segoe UI","Hiragino Kaku Gothic ProN","Meiryo",sans-serif;line-height:1.6}
  header{background:linear-gradient(135deg,#e60012,#ff5a5f);color:#fff;padding:24px 20px}
  header h1{margin:0;font-size:22px}
  header p{margin:6px 0 0;font-size:13px;opacity:.95}
  .wrap{max-width:920px;margin:0 auto;padding:18px 16px 60px}
  .controls{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:14px}
  .controls input[type=search]{flex:1 1 220px;min-width:180px;padding:9px 12px;border:1px solid var(--line);border-radius:10px;font-size:14px}
  .controls select{padding:9px 10px;border:1px solid var(--line);border-radius:10px;font-size:14px;background:#fff}
  .controls label{font-size:13px;color:var(--sub);display:flex;align-items:center;gap:5px;cursor:pointer}
  .note{font-size:12px;color:var(--sub);background:#eef2f7;border-radius:10px;padding:9px 12px;margin-bottom:14px}
  .count{font-size:13px;color:var(--sub);margin-bottom:12px}
  .card{display:flex;gap:16px;background:var(--card);border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:12px;text-decoration:none;color:inherit;transition:.15s;box-shadow:0 1px 2px rgba(20,40,60,.04)}
  .card:hover{box-shadow:0 6px 18px rgba(20,40,60,.12);transform:translateY(-1px)}
  .thumb{flex:0 0 168px;width:168px;height:94px;border-radius:8px;overflow:hidden;background:#dde3ea}
  .thumb img{width:100%;height:100%;object-fit:cover;display:block}
  .info{flex:1 1 auto;min-width:0}
  .title{font-size:16px;font-weight:700;margin:0 0 4px}
  .meta{font-size:12px;color:var(--sub);margin:0 0 6px}
  .badges{display:flex;gap:6px;flex-wrap:wrap;margin:0 0 6px}
  .badge{font-size:11px;padding:2px 8px;border-radius:999px;background:#eef2f7;color:var(--sub)}
  .badge.players{background:#e7f1ff;color:#0a6cff}
  .badge.soon{background:#fff3d6;color:#a9740a}
  .badge.low{background:#e2f7ec;color:#0a9d58}
  .desc{font-size:13px;color:#33414f;margin:6px 0 0}
  .price{font-size:15px;font-weight:700;margin-top:8px}
  .price .now{color:var(--red)}
  .price s{color:#9aa7b4;font-weight:400;font-size:13px;margin-right:6px}
  .price .off{font-size:11px;background:var(--red);color:#fff;border-radius:6px;padding:1px 6px;margin-left:6px}
  .price .free{color:#0a9d58}
  .price .tba{color:var(--sub);font-weight:500;font-size:13px}
  .histlow{font-size:12px;color:var(--low);font-weight:600;margin-top:3px}
  .histlow .lbl{color:var(--sub);font-weight:400}
  .empty{text-align:center;color:var(--sub);padding:40px}
  #more{display:block;width:100%;padding:12px;margin:6px 0 20px;font-size:14px;font-weight:700;color:var(--accent);background:#fff;border:1px solid var(--line);border-radius:12px;cursor:pointer}
  #more:hover{background:#f0f6ff}
  @media(max-width:560px){
    .card{flex-direction:column;gap:10px}
    .thumb{flex-basis:auto;width:100%;height:160px}
  }
</style>
</head>
<body>
<header>
  <h1>Nintendo Switch ゲーム一覧</h1>
  <p>「オンラインで協力」対応はチェックボックスで絞り込み / 価格出典: Nintendo 公式ソフト検索API・取得日 __DATE__</p>
</header>
<div class="wrap">
  <div class="controls">
    <input type="search" id="q" placeholder="タイトル・メーカー・ジャンルで検索">
    <select id="sort">
      <option value="pop">人気順</option>
      <option value="priceA">価格が安い順</option>
      <option value="priceD">価格が高い順</option>
      <option value="lowA">過去最安が安い順</option>
      <option value="name">名前順</option>
    </select>
    <label><input type="checkbox" id="coopOnly" checked> オンラインで協力</label>
    <label><input type="checkbox" id="saleOnly"> セール中のみ</label>
    <label><input type="checkbox" id="lowOnly"> 過去最安あり</label>
  </div>
  <div class="note">「過去最安（参考）」は <b>Deku Deals</b> の日本eShop価格追跡データに基づく参考値（「オンラインで協力」対応タイトルのみ付与）。追跡開始以降の最安値で全期間の保証ではありません。価格・割引は __DATE__ 時点。</div>
  <div class="count" id="count"></div>
  <div id="list"></div>
  <button id="more" style="display:none">さらに表示</button>
  <div class="empty" id="empty" style="display:none">該当するタイトルがありません</div>
</div>
<script>
const DATA = __DATA__;
const yen = n => '¥' + Number(n).toLocaleString('ja-JP');
function eff(r){ return (r.cur ?? r.price ?? r.dprice ?? r.pprice); }
function reg(r){ return (r.price ?? r.dprice ?? r.pprice ?? r.cur); }
function priceHtml(r){
  const e = eff(r), g = reg(r);
  if(e===null||e===undefined||e==='') return '<span class="tba">価格未定</span>';
  if(Number(e)===0) return '<span class="free">無料</span>';
  if(r.sale===1 && g && Number(g)>Number(e)){
    const off = Math.round((1-Number(e)/Number(g))*100);
    return '<s>'+yen(g)+'</s><span class="now">'+yen(e)+'</span><span class="off">'+off+'% OFF</span>';
  }
  return '<span>'+yen(e)+'</span>';
}
function lowHtml(r){
  if(r.low===null||r.low===undefined) return '';
  const e=eff(r);
  if(e!==null && e!==undefined && Number(e)<=Number(r.low))
    return '<div class="histlow"><span class="badge low">最安値圏</span> <span class="lbl">過去最安</span> '+yen(r.low)+'</div>';
  return '<div class="histlow"><span class="lbl">過去最安</span> '+yen(r.low)+'</div>';
}
const esc = s => (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function card(r){
  const soon = r.ssitu && r.ssitu!=='onsale';
  return `<a class="card" href="${esc(r.url)}" target="_blank" rel="noopener">
    <div class="thumb"><img loading="lazy" src="${esc(r.img)}" alt="${esc(r.title)}" onerror="this.style.display='none'"></div>
    <div class="info">
      <p class="title">${esc(r.title)}</p>
      <div class="badges">
        ${r.nplayer?`<span class="badge players">プレイ人数 ${esc(r.nplayer)}</span>`:''}
        ${soon?`<span class="badge soon">${esc(r.ssitu==='preorder'?'予約・近日':'未発売')}</span>`:''}
        ${r.genre?`<span class="badge">${esc(r.genre)}</span>`:''}
      </div>
      <p class="meta">${esc(r.maker)}</p>
      ${r.desc?`<p class="desc">${esc(r.desc)}</p>`:''}
      <div class="price">${priceHtml(r)}</div>
      ${lowHtml(r)}
    </div>
  </a>`;
}
const $=id=>document.getElementById(id);
const CHUNK=400;
let view=[], shown=0;
function renderMore(){
  const next=view.slice(shown, shown+CHUNK);
  $('list').insertAdjacentHTML('beforeend', next.map(card).join(''));
  shown+=next.length;
  $('more').style.display = shown<view.length ? 'block' : 'none';
  $('count').textContent='全 '+view.length.toLocaleString('ja-JP')+' タイトル'+(shown<view.length?'（'+shown.toLocaleString('ja-JP')+'件表示中）':'');
}
function render(){
  const q=$('q').value.trim().toLowerCase();
  const coop=$('coopOnly').checked;
  const sale=$('saleOnly').checked;
  const lowOnly=$('lowOnly').checked;
  const sort=$('sort').value;
  let rows=DATA.filter(r=>{
    if(coop && r.coop!==1) return false;
    if(sale && r.sale!==1) return false;
    if(lowOnly && (r.low===null||r.low===undefined)) return false;
    if(!q) return true;
    return (r.title+' '+r.maker+' '+r.genre).toLowerCase().includes(q);
  });
  if(sort==='name') rows.sort((a,b)=>(a.title||'').localeCompare(b.title||'','ja'));
  else if(sort==='priceA') rows.sort((a,b)=>(eff(a)??1e9)-(eff(b)??1e9));
  else if(sort==='priceD') rows.sort((a,b)=>(eff(b)??-1)-(eff(a)??-1));
  else if(sort==='lowA') rows.sort((a,b)=>(a.low??1e9)-(b.low??1e9));
  view=rows; shown=0;
  $('list').innerHTML='';
  $('empty').style.display=rows.length?'none':'block';
  renderMore();
}
$('more').addEventListener('click',renderMore);
['q','sort','coopOnly','saleOnly','lowOnly'].forEach(id=>$(id).addEventListener('input',render));
render();
</script>
</body>
</html>
'@
$htmlOut = $tpl.Replace('__DATA__', $recsJson).Replace('__DATE__', $today)
$indexPath = Join-Path $root 'index.html'
$htmlOut | Out-File -FilePath $indexPath -Encoding utf8
Log "生成完了: $indexPath ($([math]::Round((Get-Item $indexPath).Length/1KB)) KB)"

# ---------- 4) GitHub Pages へ公開 ----------
if($NoPush){
    Write-Host "`n[4/4] 公開はスキップ (-NoPush)。ローカル生成のみ完了。"
}
else{
    Write-Host "`n[4/4] GitHub Pages へ公開"
    $doPush = $Yes
    if(-not $Yes){
        $ans = Read-Host "  公開しますか? (Y/N)"
        $doPush = ($ans -match '^[Yy]')
    }
    if($doPush){
        Push-Location $root
        # gitのstderr警告(改行変換など)を失敗扱いにしないため、公開ブロックはContinueで実行し
        # 成否は終了コード($LASTEXITCODE)で判定する
        $prevEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try{
            git add index.html data/deku_lows.json
            git -c commit.gpgsign=false commit -q -m "Update: game list ($today)"
            # 変更が無い場合のcommit失敗(exit 1)は無視して続行(既にcommit済みの再実行など)
            $credOut = ("protocol=https`nhost=github.com`n`n" | git credential-manager get 2>$null) -split "`r?`n"
            $tok = ([string](@($credOut | Where-Object { $_ -like 'password=*' })[0]) -replace '^password=','').Trim()
            if(-not $tok){ throw "GitHubトークンを取得できませんでした。手動で 'git push' してください。" }
            $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("display-champion:$tok"))
            $branch='claude/switch-coop-games-pricing-og96e6'
            git -c http.extraheader="Authorization: Basic $b64" push origin "HEAD:$branch"
            if($LASTEXITCODE -ne 0){ throw "git push が終了コード $LASTEXITCODE で失敗しました。" }
            Write-Host "  公開しました: https://display-champion.github.io/Win/"
            Write-Host "  (反映まで1〜2分ほどかかることがあります)"
        }
        catch{ Write-Host "  公開に失敗: $($_.Exception.Message)" -ForegroundColor Yellow }
        finally{ $ErrorActionPreference = $prevEAP; Pop-Location }
    }
    else{ Write-Host "  公開をキャンセルしました。ローカルのindex.htmlは更新済みです。" }
}

Write-Host "`n完了しました。"
