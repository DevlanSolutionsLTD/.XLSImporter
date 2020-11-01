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
                $query = "INSERT INTO xls-importer (book_title, book_author, book_isbn, book_publisher, book_year_published) VALUES(?,?,?,?,?)";
                $paramType = "ssssssssssss";
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

<body class="hold-transition sidebar-mini layout-fixed layout-navbar-fixed layout-footer-fixed">
    <div class="wrapper">
        <div class="content-wrapper">
            <!-- Content Header (Page header) -->
            <section class="content-header">
                <div class="container-fluid">
                    <div class="row mb-2">
                        <div class="col-sm-6">
                            <h1>Import Students Details From .xls (Spreadsheet) File</h1>
                        </div>
                        <div class="col-sm-6">
                            <ol class="breadcrumb float-sm-right">
                                <li class="breadcrumb-item"><a href="#">Home</a></li>
                                <li class="breadcrumb-item"><a href="dashboard.php">Dashboard</a></li>
                                <li class="breadcrumb-item"><a href="manage_students.php">Students</a></li>
                                <li class="breadcrumb-item active">Import</li>
                            </ol>
                        </div>
                    </div>
                </div><!-- /.container-fluid -->
            </section>

            <!-- Main content -->
            <section class="content">
                <div class="container-fluid">
                    <div class="col-md-12">
                        <!-- general form elements -->
                        <div class="card card-primary">
                            <div class="card-header">
                                <h3 class="card-title text-danger">*Beta Module</h3>
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
                        </div>
                    </div>
                </div>
            </section>
            <!-- /.content -->
        </div>
    </div>
</body>

</html>
