<#
概要：
指定したSharePointサイトのライブラリに、指定したディレクトリ配下のファイルをアップロードする
実行形式：
./spLibUpload.ps1 -url [サイトのURL] -library [アップロード先ライブラリ名] -username [ユーザー名] -password [パスワード] -path [対象のファイルのあるディレクトリパス]
#>
#
# SharePoint Online Automation – O365 – Upload files to document library using PowerShell CSOM
# https://global-sharepoint.com/powershell/upload-files-to-sharepoint-online-document-library-using-powershell-csom/


# 引数取得(サイトURL、ユーザー名、パスワード、登録するファイルのあるフォルダ)
Param(
    [parameter(mandatory=$true)][String]$url,
    [parameter(mandatory=$true)][String]$library,
    [parameter(mandatory=$true)][String]$username,
    [parameter(mandatory=$true)][String]$password,
    [parameter(mandatory=$true)][String]$path
) 

# CSOMライブラリの読み込み
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") > $null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") > $null
#Add-Type -Path "C:\Users\circleci\project\Microsoft.SharePointOnline.CSOM.16.1.20211.12000\lib\netstandard2.0\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Users\circleci\project\Microsoft.SharePointOnline.CSOM.16.1.20211.12000\lib\netstandard2.0\Microsoft.SharePoint.Client.Runtime.dll"
#Import-Module "C:\Users\circleci\project\Microsoft.SharePointOnline.CSOM.16.1.20211.12000\lib\net45\Microsoft.SharePoint.Client.dll"
#Import-Module "C:\Users\circleci\project\Microsoft.SharePointOnline.CSOM.16.1.20211.12000\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"

# ログイン処理
$securepassword = ConvertTo-SecureString $password -AsPlainText -Force
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($url)
$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securepassword)

# "For Each" loop will upload all of the files one by one onto the destination using the UploadFile method
$filesCollectionInSourceDirectory=Get-ChildItem $path -File   

ForEach ($oneFile in $filesCollectionInSourceDirectory) {
      
    $list = $ctx.Web.Lists.GetByTitle($library)
    $ctx.Load($list)
    $ctx.ExecuteQuery()     
   
    $SourceFilePath=$oneFile.FullName
    $targetFilePath=$url+"/"+"$library"+"/"+$oneFile

    $fileOpenStream = New-Object IO.FileStream($SourceFilePath, [System.IO.FileMode]::Open)  
    $fileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation  
    $fileCreationInfo.Overwrite = $true  
    $fileCreationInfo.ContentStream = $fileOpenStream  
    $fileCreationInfo.URL = $oneFile
    $uploadFileInfo = $list.RootFolder.Files.Add($FileCreationInfo)  
    $ctx.Load($uploadFileInfo)  
    $ctx.ExecuteQuery() 
     
    Write-host -f Green "File '$SourceFilePath' has been uploaded to '$targetFilePath' successfully!"
}