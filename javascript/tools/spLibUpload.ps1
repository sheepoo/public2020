<#
概要：
指定したSharePointサイトのライブラリに、指定したディレクトリ配下のファイルをアップロードする
実行形式：
./spLibUpload.ps1 -url [サイトのURL] -library [アップロード先ライブラリ名] -username [ユーザー名] -password [パスワード] -path [対象のファイルのあるディレクトリのフルパス]
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
#Add-Type -Path "/home/circleci/.local/share/powershell/Modules/SharePointOnline.CSOM/1.0.5/Microsoft.SharePoint.Client.dll"
#Add-Type -Path "/home/circleci/.local/share/powershell/Modules/SharePointOnline.CSOM/1.0.5/Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "/home/circleci/.local/share/powershell/Modules/SharePointOnline.CSOM/1.0.5/Microsoft.SharePoint.Client.Portable.dll"
Add-Type -Path "/home/circleci/.local/share/powershell/Modules/SharePointOnline.CSOM/1.0.5/Microsoft.SharePoint.Client.Runtime.Portable.dll"
Add-Type -Path "/home/circleci/.local/share/powershell/Modules/SharePointOnline.CSOM/1.0.5/Microsoft.SharePoint.Client.Runtime.Windows.dll"

# ログイン処理
$securepassword = ConvertTo-SecureString $password -AsPlainText -Force
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($url)
$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securepassword)

# "For Each" loop will upload all of the files one by one onto the destination using the UploadFile method
$filesCollectionInSourceDirectory=Get-ChildItem $path -File   

ForEach ($oneFile in $filesCollectionInSourceDirectory) {

    try {   
            #リスト一覧を取得
            $ctx.Load($ctx.Web.lists)
            $ctx.ExecuteQuery()
            foreach($objList in $ctx.Web.lists)
            {
                $listTitle = $objList.Title
                Write-Host "リスト ${listTitle} をエクスポートします"
            }

            Write-Host "ここからライブラリ処理"
            $list = $ctx.Web.Lists.GetByTitle($library)
            $ctx.Load($list)
            $ctx.ExecuteQuery()     
            Write-Host "ここからアップロード"

            $SourceFilePath=$oneFile.FullName
            $targetFilePath=$url+"/"+"$library"+"/"+$oneFile.Name
            
            $fileOpenStream = New-Object IO.FileStream($SourceFilePath, [System.IO.FileMode]::Open)  
            $fileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation  
            $fileCreationInfo.Overwrite = $true  
            $fileCreationInfo.ContentStream = $fileOpenStream  
            $fileCreationInfo.URL = $oneFile
            $uploadFileInfo = $list.RootFolder.Files.Add($FileCreationInfo)  
            $ctx.Load($uploadFileInfo)  
            $ctx.ExecuteQuery() 
            
            Write-host "File '$SourceFilePath' has been uploaded to '$targetFilePath' successfully!"
    }
    catch {
        Write-Error ($_.Exception)
        Write-Host "Upload error : File '$SourceFilePath'"
        exit 1
    }
}