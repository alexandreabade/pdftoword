<?php
// Caminho do arquivo no servidor
$filePath = __DIR__ . '/temp/arquivo_convertido.docx';

// Verifica se o arquivo existe
if (file_exists($filePath)) {
    // Define cabeçalhos para download
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="' . basename($filePath) . '"');
    header('Content-Length: ' . filesize($filePath));

    // Lê o conteúdo do arquivo e envia para o navegador
    readfile($filePath);

    // Encerra o script após o download
    exit;
} else {
    // Se o arquivo não existir
    echo "O arquivo não existe.";
}
?>
