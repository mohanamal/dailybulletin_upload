<?php
// error_reporting(0);
require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;


if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}


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


$client = getClient();
$service = new Google_Service_Sheets($client);
$spreadsheetId = '15Qfe9FQwUSUtRTS0RPkYg0VHdNztiUFmwvtoWjBjdwM';
$range = "testconnection!A1:E";
$response = $service->spreadsheets_values->get($spreadsheetId, $range);
$values = $response->getValues();
if (empty($values)) {
    print "No data found.\n";
} else {
    print "Name, Major:\n";
    foreach ($values as $row) {
        // Print columns A and E, which correspond to indices 0 and 4.
        printf("%s, %s\n", $row[0], $row[1]);
    }
}




// $pptReader = IOFactory::createReader('PowerPoint2007');
// $oPHPPresentation = $pptReader->load('tt.pptx');

// $dat = $oPHPPresentation->getAllSlides();
// foreach ($dat as $key => $oSlide) {
// 	if ($key != 3)
// 		continue;
// 	foreach ($oSlide->getShapeCollection() as $oShape) {
// 		if ($oShape instanceof RichText) {
// 			// foreach ($oShape->getParagraphs() as $oParagraph) {
// 			// 	foreach ($oParagraph->getRichTextElements() as $oRichText) {
// 			// 		echo "text=" . $oRichText->getText();
// 			// 		echo "<br/>";
// 			// 	}
// 			// }
// 		} else if ($oShape instanceof Table) {
// 			foreach ($oShape->getRows() as $key => $value) {
// 				foreach ($value->getCells() as $cell) {
// 					foreach ($cell->getParagraphs() as $oParagraph) {
// 						foreach ($oParagraph->getRichTextElements() as $oRichText) {
// 							echo "text=" . $oRichText->getText();
// 							echo "<br/>";
// 						}
// 					}
// 				}
// 			}
// 		}
// 	}
// }