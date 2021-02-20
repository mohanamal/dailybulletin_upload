<?php
ini_set('max_execution_time', '0'); // for infinite time of execution 
// error_reporting(0);
require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
// if (php_sapi_name() != 'cli') {
//     throw new Exception('This application must be run on the command line.');
// }

/**
 * Returns an authorized API client.
 * @return Google_Client the authorized client object
 */
function getClient()
{
    $client = new Google_Client();
    $client->setRedirectUri("urn:ietf:wg:oauth:2.0:oob");
    $client->setApplicationName('Web client 2');
    $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
    $client->setAuthConfig(__DIR__ . '/credentials.json');
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');
    // Load previously authorized token from a file, if it exists.
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    $tokenPath = __DIR__ . '/token.json';
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
        $client->setAccessToken($accessToken);
    }

    // If there is no previous token or it's expired.
    if ($client->isAccessTokenExpired()) {
        // Refresh the token if possible, else fetch a new one.
        if ($client->getRefreshToken()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
        } else {
            // Request authorization from the user.
            $authUrl = $client->createAuthUrl();
            printf("Open the following link in your browser:\n%s\n", $authUrl);
            print 'Enter verification code: ';
            $authCode = trim(fgets(STDIN));

            // Exchange authorization code for an access token.
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
            $client->setAccessToken($accessToken);

            // Check to see if there was an error.
            if (array_key_exists('error', $accessToken)) {
                throw new Exception(join(', ', $accessToken));
            }
        }
        // Save the token to a file.
        if (!file_exists(dirname($tokenPath))) {
            mkdir(dirname($tokenPath), 0700, true);
        }
        file_put_contents($tokenPath, json_encode($client->getAccessToken()));
    }
    return $client;
}
function uplodToGoogle($datArr, $sheetName)
{
    echo  "<br/>" . $sheetName;
    // Get the API client and construct the service object.
    $client = getClient();
    $service = new Google_Service_Sheets($client);
    $spreadsheetId = '15Qfe9FQwUSUtRTS0RPkYg0VHdNztiUFmwvtoWjBjdwM';
    // $range = "COVIDHospital";
    $range = $sheetName;
    $values = $datArr;
    $body = new Google_Service_Sheets_ValueRange([
        'values' => $values
    ]);
    $conf = ["valueInputOption" => "RAW"];
    $service->spreadsheets_values->append($spreadsheetId, $range, $body, $conf);
}

function getTableDetails()
{
    $source = 'tt.docx';
    $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
    $sections = $phpWord->getSections();
    $i = 0;
    $tableHead = 'Start';
    $tableArray = array();

    foreach ($sections as $key => $sec) {
        foreach ($sec->getElements() as $el) {
            $oldTbleHead = ($tableHead != '') ? $tableHead : $oldTbleHead;
            if ($el instanceof PhpOffice\PhpWord\Element\Table) {
                $tableHead = ($tableHead != '') ? $tableHead : $oldTbleHead;
                $inarray = array(
                    'tblHead' => preg_replace('/\s+/', '', strtolower($tableHead)),
                    'tblId' => $i++,
                );
                array_push($tableArray, $inarray);
                $tableHead = '';
            } else {
                if (get_class($el) === 'PhpOffice\PhpWord\Element\Text') {
                    // $tableHead .= (preg_replace('/\s+/', '', $el->getText()) != '') ? $el->getText() : $tableHead;
                    echo "<br/>dddddd" .  $outText =  $el->getText();
                    $tableHead .=  $outText;
                } else if (get_class($el) === 'PhpOffice\PhpWord\Element\TextRun') {
                    $tableHead = '';
                    foreach ($el->getElements() as $text) {
                        // $outText = preg_replace('/\s+/', '', $text->getText());
                        $outText =  $text->getText();
                        $tableHead .=  $outText;
                    }
                    // echo "<br/>" . $tableHead;
                    // $tableHead = (($tableHead != '')) ? $tableHead : $oldTbleHead;
                } //
            } //return $finalArrayTemp;

            // echo "<br/>second:" . $tableHead;
        } //foreach ($sec->getElements() as $el)
    } //foreach ($sections as $key => $sec) 
    return ($tableArray);
}

function getTableData($date, $tblDetails)
{
    $source = 'tt.docx';
    $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
    $sections = $phpWord->getSections();
    $i = 0;
    $tableArray = array();
    foreach ($sections as $key => $sec) {
        foreach ($sec->getElements() as $el) {
            if ($el instanceof PhpOffice\PhpWord\Element\Table) {
                if ($tblDetails['tblId'] != $i++)
                    continue;
                $finalArrayTemp = array();
                foreach ($el->getRows() as $row) {
                    $rowArray = array($date);
                    foreach ($row->getCells() as $cell) {
                        $celements = $cell->getElements();
                        $inVal = '';
                        foreach ($celements as $celem) {
                            if (get_class($celem) === 'PhpOffice\PhpWord\Element\Text') {
                                $inVal .=  $celem->getText();
                            } else if (get_class($celem) === 'PhpOffice\PhpWord\Element\TextRun') {
                                foreach ($celem->getElements() as $text) {
                                    $inVal .= $text->getText();
                                }
                            }
                        } //
                        array_push($rowArray, $inVal);
                    }
                    array_push($finalArrayTemp, $rowArray);
                }
                return $finalArrayTemp;
            } else {
            }
        } //foreach ($sec->getElements() as $el)
    } //foreach ($sections as $key => $sec) 
    return ($tableArray);
}
function loadUiData()
{
    // $date = '03-02-2021';
    $tableDet = getTableDetails(); //fetch all table details
    $annexure1Array = array();
    echo "<pre>";
    print_r($tableDet);
    echo "</pre>";
    foreach ($tableDet as $key => $value) {
        $tableHead = substr($value['tblHead'], 0, 10);
        switch ($tableHead) {
            case 'table2.sum':
                $sheetName = 'table2.summaryofnewcovid-19cases';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                break;
            case 'table3.cum':
                // echo "sdsf";
                $sheetName = 'table3.cumulativesummary';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                break;
            case 'table4.dis':
                $sheetName = 'table4.summaryOfCumulativeActiveNewPositiveCases';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                break;
            case 'table5.sum':
                $sheetName = 'tbl5SummaryOfContactorTravelHistoryCasesSinceMay4t';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                break;
            case 'table6.sum':
                $sheetName = 'table6.SummaryOfcontact/travelHistoryOfNewCases';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'table7.det':
                $sheetName = 'table7.DetailsOfPositiveDeathsInTheLast24hours';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'table8.sum':
                $sheetName = 'table8.SummaryOfCriticalPatients';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'table11.di':
                $sheetName = 'table11.DistributionOfQuarantine&isolation';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'table12.su':
            case 'summaryoft':
                $sheetName = 'table12.SummaryofTravelSurveillance';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'table.13.p':
                $sheetName = 'table.13.psychosocialsupportprovided';
                $tblData = getTableData($date, $value);
                uplodToGoogle($tblData, $sheetName);
                // echo "<pre>";
                // print_r($tblData);
                // echo "</pre>";
                // die();
                break;
            case 'annexure-1':
                $sheetName = 'annexure-1:listofhotspots';
                $tblData = getTableData($date, $value);
                $annexure1Array = array_merge($annexure1Array, $tblData);
                break;
            case 'annexure-2':
            case 'annexure-3':
            case 'table1.sum':
            case 'table9.cum':
            case 'table10.de':
                $sheetName = 'notNeeded';
                break;
            default:
                $sheetName = "unknown";
                // echo "<br/>" . $tableHead;
                $tblData = getTableData($date, $value);
                array_push($tblData, array($tableHead));
                uplodToGoogle($tblData, $sheetName);
                break;
        }
    }
    $splitedArray = splitArrayIntoThree($annexure1Array); //tofind the table of three contents
    //hotspotsDeleted
    uplodToGoogle($splitedArray['hotspots_deleted'], 'hotspotsDeleted');
    $hotspots_deleted_count =  getCountsOfHotSpot($splitedArray['hotspots_deleted'], $date);
    uplodToGoogle($hotspots_deleted_count, 'hotspotsDeletedCount');
    //hotspotsAdded
    uplodToGoogle($splitedArray['hotspots_added'], 'hotspotsAdded');
    $hotspots_added_count =  getCountsOfHotSpot($splitedArray['hotspots_added'], $date);
    uplodToGoogle($hotspots_added_count, 'hotspotsAddedCount');
    //hotspots
    uplodToGoogle($splitedArray['hotspots'], 'hotspots');
    $hotspots =  getCountsOfHotSpot($splitedArray['hotspots'], $date);
    uplodToGoogle($hotspots, 'hotspotsCount');

    // echo "<pre>";
    // print_r($splitedArray['hotspots']);
    // echo "</pre>";
    // die();
    die();
}

function splitArrayIntoThree($inData)
{
    $secondTableStart =  searchForId('LSGs added on', $inData);
    $firstSplit =  sliceArray($inData, $secondTableStart);
    $thirdTableStart = searchForId('LSGs deleted due to no containment zone', $firstSplit['second']);
    $secondSplit =  sliceArray($firstSplit['second'], $thirdTableStart);
    $out = array(
        'hotspots' => $firstSplit['first'],
        'hotspots_added' => $secondSplit['first'],
        'hotspots_deleted' => $secondSplit['second'],
    );
    return $out;
}
function searchForId($id, $array)
{
    foreach ($array as $key => $val) {
        if (strpos($val[1], $id) !== false) {
            return $key;
        }
    }
    return null;
}

function sliceArray($inArray, $splitKey)
{
    $firstArray = array();
    $secondArray = array();
    foreach ($inArray as $key => $value) {
        if ($key < $splitKey)
            array_push($firstArray, $value);
        else
            array_push($secondArray, $value);
    }
    $out = array(
        'first' => $firstArray,
        'second' => $secondArray
    );
    return $out;
}

function getCountsOfHotSpot($inArray, $date)
{
    $outCompailed = array();
    $out = array();
    foreach ($inArray as $key => $value) {
        if (is_numeric($value[1])) {
            $countofHot = count(explode(",", $value[4]));
            if (!isset($out[$value[2]])) {
                $out[$value[2]] = array(
                    $date, $value[2], 1, $countofHot
                );
            } else {
                $out[$value[2]][2] = $out[$value[2]][2] + 1;
                $out[$value[2]][3] += $countofHot;
            }
        } else
            continue;
    }

    foreach ($out as $key => $value) {
        array_push($outCompailed, $value);
    }
    return $outCompailed;
}

loadUiData();
