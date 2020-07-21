<#
概要：
指定したSharePointサイトのライブラリに、指定したディレクトリ配下のファイルをアップロードする
実行形式：
./spLibUpload.ps1 -url [サイトのURL] -library [アップロード先ライブラリ名] -username [ユーザー名] -password [パスワード] -path [対象のファイルのあるディレクトリのフルパス]
#>


# 引数取得(サイトURL、ユーザー名、パスワード、登録するファイルのあるフォルダ)
Param(
    [parameter(mandatory=$true)][String]$url,
    [parameter(mandatory=$true)][String]$library,
    [parameter(mandatory=$true)][String]$username,
    [parameter(mandatory=$true)][String]$password,
    [parameter(mandatory=$true)][String]$path
) 

# ログイン処理
$securePassword = convertto-securestring -String $password -AsPlainText -Force
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $securePassword
Connect-PnPOnline -Url $url -Credentials $cred

# "For Each" loop will upload all of the files one by one onto the destination using the UploadFile method
$filesCollectionInSourceDirectory=Get-ChildItem $path -File   

ForEach ($oneFile in $filesCollectionInSourceDirectory) {
    try {   
            $SourceFilePath=$oneFile.FullName
            $pnplibrary = "/"+$library
            $file = Add-PnPFile -Path $SourceFilePath -Folder "$pnplibrary" 
            Write-host "File '$SourceFilePath' has been uploaded to '$pnplibrary' successfully!"
    }
    catch {
        Write-Error ($_.Exception)
        Write-Host "Upload error : File $SourceFilePath"
        exit 1
    }
}