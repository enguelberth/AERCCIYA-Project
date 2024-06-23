<?php
session_start();
include("../model/config.php");

if (!isset($_SESSION['valid'])) {
    header("Location: ../index.php");
    exit();
}

require '../vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;

if (isset($_GET['id'])) {
    $id_extraordinario = $_GET['id'];

    try {
        $pdo = new PDO('mysql:host=localhost;dbname=comite_sena', 'root', '');
        $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

        // Obtener datos del comité
        $stmt = $pdo->prepare("SELECT * FROM comite_extraordinario WHERE ID_extraordinario = :id");
        $stmt->execute(['id' => $id_extraordinario]);
        $comite = $stmt->fetch(PDO::FETCH_ASSOC);

        if ($comite) {
            // Crear el documento a partir de la plantilla
            $templateProcessor = new TemplateProcessor('../static/template.docx');

            // Rellenar la plantilla con los datos del comité
            $templateProcessor->setValue('Acta_Num', $comite['Acta_Num']);
            $templateProcessor->setValue('Nombre', $comite['Nombre']);
            $templateProcessor->setValue('Fecha', $comite['Fecha']);
            $templateProcessor->setValue('Hora_inicio', $comite['Hora_inicio']);
            $templateProcessor->setValue('Hora_fin', $comite['Hora_fin']);
            $templateProcessor->setValue('Agendas', $comite['Agendas']);
            $templateProcessor->setValue('Objetivo', $comite['Objetivo']);
            $templateProcessor->setValue('Desarrollo', $comite['Desarrollo']);
            $templateProcessor->setValue('Responsable', $comite['Responsable']);

            // Guardar el documento
            $fileName = 'comite_extraordinario_' . $id_extraordinario . '.docx';
            $templateProcessor->saveAs('../uploads/' . $fileName);

            // Descargar el archivo
            header('Content-Description: File Transfer');
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Disposition: attachment; filename="' . basename($fileName) . '"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');
            header('Content-Length: ' . filesize('../uploads/' . $fileName));
            readfile('../uploads/' . $fileName);
            exit;
        } else {
            echo "Comité no encontrado.";
        }
    } catch (PDOException $e) {
        echo "Error de conexión: " . $e->getMessage();
    }
} else {
    echo "ID del comité no especificado.";
}
?>