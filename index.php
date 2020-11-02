<?php

use DevlanInc\DataSource;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

require_once('DataSource.php');
$db = new DataSource();
$conn = $db->getConnection();
require_once('vendor/autoload.php');

if (isset($_POST["upload"])) {

    $allowedFileType = [
        'application/vnd.ms-excel',
        'text/xls',
        'text/xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];

    if (in_array($_FILES["file"]["type"], $allowedFileType)) {

        $targetPath = 'uploads/' . $_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $targetPath);

        $Reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

        $spreadSheet = $Reader->load($targetPath);
        $excelSheet = $spreadSheet->getActiveSheet();
        $spreadSheetAry = $excelSheet->toArray();
        $sheetCount = count($spreadSheetAry);

        for ($i = 0; $i <= $sheetCount; $i++) {

            $book_title = "";
            if (isset($spreadSheetAry[$i][1])) {
                $book_title = mysqli_real_escape_string($conn, $spreadSheetAry[$i][1]);
            }
            $book_isbn = "";
            if (isset($spreadSheetAry[$i][2])) {
                $book_isbn = mysqli_real_escape_string($conn, $spreadSheetAry[$i][2]);
            }

            $book_author = "";
            if (isset($spreadSheetAry[$i][3])) {
                $book_author = mysqli_real_escape_string($conn, $spreadSheetAry[$i][3]);
            }

            $book_publisher = "";
            if (isset($spreadSheetAry[$i][4])) {
                $book_publisher = mysqli_real_escape_string($conn, $spreadSheetAry[$i][4]);
            }

            $book_year_published = "";
            if (isset($spreadSheetAry[$i][5])) {
                $book_year_published = mysqli_real_escape_string($conn, $spreadSheetAry[$i][5]);
            }


            if (!empty($book_title) || !empty($book_author) || !empty($book_isbn) || !empty($book_year_published) || !empty($book_publisher)) {
                $query = "INSERT INTO xls_importer_demo (book_title, book_author, book_isbn, book_publisher, book_year_published) VALUES(?,?,?,?,?)";
                $paramType = "sssss";
                $paramArray = array(
                    $book_title,
                    $book_author,
                    $book_isbn,
                    $book_publisher,
                    $book_year_published
                );
                $insertId = $db->insert($query, $paramType, $paramArray);
                if (!empty($insertId)) {
                    $success = "Excel Data Imported into the Database";
                } else {
                    $success = "Errors Encountered When Importing Data";
                }
            }
        }
    } else {
        $info = "Invalid File Type. Upload Excel File.";
    }
}

?>
<!doctype html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta name="description" content="">
    <meta name="author" content="Mark Otto, Jacob Thornton, and Bootstrap contributors">
    <meta name="generator" content="Jekyll v4.0.1">
    <title>.XLS Importer Librayr</title>

    <link rel="canonical" href="https://getbootstrap.com/docs/4.5/examples/dashboard/">

    <!-- Dashboard core CSS -->
    <link href="Assets/dashboard.css" rel="stylesheet">
    <!-- Bootstrap Core CSS -->
    <link href="Assets/bootstrap.css" rel="stylesheet">

    <style>
        .bd-placeholder-img {
            font-size: 1.125rem;
            text-anchor: middle;
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
        }

        @media (min-width: 768px) {
            .bd-placeholder-img-lg {
                font-size: 3.5rem;
            }
        }
    </style>
    <!-- Custom styles for this template -->
    <link href="dashboard.css" rel="stylesheet">
</head>

<body>
    <nav class="navbar navbar-dark sticky-top bg-dark flex-md-nowrap p-0 shadow">
        <a class="navbar-brand col-md-3 col-lg-2 mr-0 px-3" href="#">.XLS Importer Library</a>
        <button class="navbar-toggler position-absolute d-md-none collapsed" type="button" data-toggle="collapse" data-target="#sidebarMenu" aria-controls="sidebarMenu" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>
        <input class="form-control form-control-dark w-100" type="text" placeholder="Search" aria-label="Search">
    </nav>

    <div class="container-fluid">
        <div class="row">
            <main role="main" class="col-md-12 ml-sm-auto col-lg-12 px-md-4">
                <div class="d-flex justify-content-between flex-wrap flex-md-nowrap align-items-center pt-3 pb-2 mb-3 border-bottom">
                    <h1 class="h2">.XLS Importer</h1>
                </div>
                <div class="card card-primary">
                    <div class="card-header">
                        <h3 class="card-title text-danger">Import An Excel Worksheet To MariaDB Using PHP</h3>
                    </div>
                    <!-- form start -->
                    <form method="post" enctype="multipart/form-data" role="form">
                        <div class="card-body">
                            <div class="row">
                                <div class="form-group col-md-12">
                                    <label for="exampleInputFile">Select File</label>
                                    <div class="input-group">
                                        <div class="custom-file">
                                            <input required name="file" accept=".xls,.xlsx" type="file" class="custom-file-input" id="exampleInputFile">
                                            <label class="custom-file-label" for="exampleInputFile">Choose file</label>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="card-footer">
                            <button type="submit" name="upload" class="btn btn-primary">Upload File</button>
                        </div>
                    </form>
                    <hr>
                    <h2>Book Details</h2>
                    <div class="table-responsive">
                        <?php
                        $sqlSelect = "SELECT * FROM xls_importer_demo";
                        $result = $db->select($sqlSelect);
                        if (!empty($result)) { {
                        ?>
                                <table class="table table-striped table-sm">
                                    <thead>
                                        <tr>
                                            <th>Title</th>
                                            <th>ISBN</th>
                                            <th>Author</th>
                                            <th>Publisher</th>
                                            <th>Year Published</th>
                                            <th>Uploaded At</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <?php
                                        foreach ($result as $row) {
                                        ?>
                                            <tr>
                                                <td><?php echo $row['book_title']; ?></td>
                                                <td><?php echo $row['book_isbn']; ?></td>
                                                <td><?php echo $row['book_author']; ?></td>
                                                <td><?php echo $row['book_publisher']; ?></td>
                                                <td><?php echo $row['book_year_published']; ?></td>
                                                <td><?php echo $row['created_at']; ?></td>
                                            </tr>
                                        <?php
                                        }
                                        ?>
                                    </tbody>
                                </table>
                        <?php
                            }
                        }
                        ?>
                    </div>
                </div>
            </main>
        </div>
    </div>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script>
        window.jQuery || document.write('<script src="../assets/js/vendor/jquery.slim.min.js"><\/script>')
    </script>
    <script src="../assets/dist/js/bootstrap.bundle.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/feather-icons/4.9.0/feather.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.7.3/Chart.min.js"></script>
    <script src="Assets/dashboard.js"></script>
</body>

</html>