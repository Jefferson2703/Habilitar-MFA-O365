#------- CONECTA NA CONTA DO OFFICE 365---------#
Install-Module MSOnline
Connect-MsolService

#------- CONVERTE EXCEL PARA CSV---------#
Function ConvertExcelCsv ($roadFile, $excelFileName, $csvLoc)
{
    $excelFile = $roadFile + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $excelFileName # + "_" + $ws.Name
        $ws.SaveAs($csvLoc + $n + ".csv", 6)
    }
    $E.Quit()
}

#------- CAMINHO DA PASTA ONDE CONTEM O XLSX ---------#
$caminho = Read-Host 'Qual o caminho completo do arquivo ?'

#------- REMOVE ESPAÇOS EXTRAS NO INICIO E NO FIM DO CAMINHO ---------#
$caminho = $caminho.trim()

#------- VALIDA DE O CAMINHO EXISTE  ---------#
while ((Test-Path -Path $caminho ) -ne 'True'){
    $caminho = Read-Host 'Caminho inválido, favor digite o local do arquivo novamente'
}

#------- AJUSTE NO CAMINHO ---------#
if ($caminho.Substring($caminho.length - 1) -eq "\"){
    $caminho = $caminho
}elseif ($caminho.Substring($caminho.length - 1) -ne "\"){
    $caminho = $caminho + "\"
}

#------- NOME DO ARQUIVO SEM EXTENSÃO ---------#
$nomeArquivo = Read-Host 'Qual o nome do arquivo?'

#------- VALIDA SE O ARQUIVO EXISTE ---------#
while ((Test-Path -Path ($nomeArquivo + ".xlsx") ) -ne 'True'){
    $nomeArquivo = Read-Host 'Arquivo não encontrado, digite novamente o nome do arquivo'
}

#------- CHAMA A FUNÇÃO PARA CONVERTER O ARQUIVO PARA CSV ---------#
ConvertExcelCsv -roadFile $caminho -excelFileName $nomeArquivo -csvLoc $caminho

#------- CAMINHO PARA BUSCAR O CSV---------#
$arquivo = $caminho + $nomeArquivo + ".csv"

#------- lista de usuários para atualizar em massa---------#
$users = import-CSV $arquivo 

#------- Habilitando MFA para usuários ---------#
foreach ($user in $users)
{
    $email = $user.email

    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.RelyingParty = "*"
    $st.State = "Enabled"
    $sta = @($st)
    Set-MsolUser -UserPrincipalName $email -StrongAuthenticationRequirements $sta

    Write-Warning "O e-mail $email foi habilitado MFA" #MENSAGEM PROMPT 
}

#------- Remove CSV gerado temporario ---------#
Remove-Item $arquivo -force -recurse


Write-Warning "---------  Exportando relatório MFA ------ $_" #MENSAGEM PROMPT 

#------- Export usuários que estão habilitados MFA ---------#
Get-MsolUser -All | Select-Object @{N='UserPrincipalName';E={$_.UserPrincipalName}},
@{N='MFA Status';E={if ($_.StrongAuthenticationRequirements.State){$_.StrongAuthenticationRequirements.State} else {"Disabled"}}},
@{N='MFA Methods';E={$_.StrongAuthenticationMethods.methodtype}} | Export-Csv -Path $caminho\MFA_Report.csv -NoTypeInformation

Pause