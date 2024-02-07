<?php



if (isset($_FILES['pdfFile']) && $_FILES['pdfFile']['error'] === UPLOAD_ERR_OK) {

    $wordFilePath = __DIR__ . '\temp\arquivo_convertido.docx';
    // Caminho do arquivo temporário
    $tmpPdfFilePath = $_FILES['pdfFile']['tmp_name'];
    $tempFileName = $_FILES['pdfFile']['tmp_name'];
    $originalFileName = $_FILES['pdfFile']['name'];

    // Diretório temporário para salvar o arquivo
    $tempDir = __DIR__ . '/temp/';

    // Gera um nome único para o arquivo
    $uniqueFileName = $tempDir . uniqid() . '_' . $originalFileName;

    // Move o arquivo para o diretório temporário
    if (move_uploaded_file($tempFileName, $uniqueFileName)) {
      //  echo 'Arquivo enviado com sucesso para: ' . $uniqueFileName;

    } else {
      //  echo 'Erro ao mover o arquivo para o diretório temporário.';
    }
    
    $word = new COM("Word.Application") or die("Não foi possível iniciar o Word.");
    $word->Visible = true;
    $document = $word->Documents->Open($uniqueFileName);
    // Realize operações no documento, se necessário
 

    $word->ActiveDocument->SaveAs($wordFilePath);
    $word->Quit();
    
    if (file_exists($uniqueFileName)) {
        // Tenta excluir o arquivo
        if (unlink($uniqueFileName)) {
           // echo 'Arquivo excluído com sucesso.';
        } else {
           // echo 'Erro ao tentar excluir o arquivo.';
        }
    } 
   
    echo '<div class="alert alert-success" role="alert">Conversão bem-sucedida! <a href="download.php" target="_blank">Clique aqui para baixar o arquivo</a></div>';
} else {
    // Se nenhum arquivo foi enviado ou ocorreu algum erro
    //echo '<div class="alert alert-danger" role="alert">Erro ao processar o arquivo PDF. Certifique-se de ter selecionado um arquivo.</div>';
}
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor PDF para Word</title>
    <!-- Adicione o link para o Bootstrap CSS -->
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
</head>

<body>
    <div class="container mt-5">
        <h2 class="mb-4">Conversor PDF para Word</h2>
        <form method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label for="pdfFile">Selecione um arquivo PDF:</label>
                <input type="file" class="form-control-file" name="pdfFile" id="pdfFile" accept=".pdf" required>
            </div>
            <button type="submit" class="btn btn-primary">Converter</button>
        </form>
    </div>

    <!-- Adicione o link para o Bootstrap JS e o jQuery -->
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.2/dist/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</body>

</html>
