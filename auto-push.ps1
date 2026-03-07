cd "C:\Users\iNFORMARK Loja\Desktop\iphone-inteligencia\Bot"

# gera o html atualizado
powershell -ExecutionPolicy Bypass -File ".\gerar-html.ps1"

# adiciona alterações
git add .

# cria commit automático com data
$data = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
git commit -m "atualizacao automatica $data"

# envia para github
git push