# AutoOrderFax

## Process Flow
1. CreateFiles
1. ftpサーバに接続
1. CleanRemoteFiles
1. TransferFiles
1. ftpサーバから切断
1. DeleteLocalFiles
1. DeleteLogFiles

## CreateFiles
   - CreateOrderSlip  
    FAX送信対象のデータがあったらローカルにpdfとreq作成

## ftpサーバに接続
```
   <FTPServer>  
      <Host>auto.lcloud.jp</Host>  
      <Port>22</Port>  
      <User>F000540001</User>  
      <Password>J9rtb6UR</Password>   
   </FTPServer>  
  
   <ProxyServer>  
      <Host>cs-sunnet</Host>  
      <Port>8080</Port>  
      <User></User>  
      <Password></Password>  
   </ProxyServer>  
```

## CleanRemoteFiles
   - リモートのpdfとreq 削除  
   - リモートのlock削除  
    想定外のご送信を防ぐため、lockは最後に消す

## TransferFiles
   - ローカルにlock作成  
   - リモートにlock転送  
   - リモートにpdfとreq転送  
   - リモートのlock削除  
   **ここでFAX送信**
   - リモートのファイルが削除される  
 
## ftpサーバから切断

## DeleteLocalFiles
   - ローカルの全てのファイル削除  

## DeleteLogFiles
   - 指定期間を過ぎているlogファイル削除  


