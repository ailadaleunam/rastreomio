<?php
session_start();
define("file_active_user_id" , "ailadaleunam_active_user_id.txt");
define("file_access_token" , "ailadaleunam_access_token.txt");
define("file_user_data" , "ailadaleunam_user_data.txt");
define("file_user_log", "ailadaleunam_log.txt");
define("file_auth_config" , "ailadaleunam_auth_config.txt");
define("CLIENTE_ID" , "1f502fb9-434e-402f-bc4b-e16af0a46245");
define("SECRETO_ID" , "iJJI3762=fscsrmOZBV4]|%");
define("CLIENTE_URL" , "https://rastreomio.vercel.app/api/ailadaleunam.php");


init();

if(isset($_GET["check_email"])) {
    check_email();
}

if(token()) {
    echo "<a href='".$_SESSION["redirect_uri"]."''>Home</a>";
    echo " || <a href='".$_SESSION["redirect_uri"]."?refresh_token=true'>Refresh token</a>";
    echo " || <a href='".$_SESSION["redirect_uri"]."?profile=true'>Profile</a>";
    echo " || <a href='".$_SESSION["redirect_uri"]."?list_log=true'>Log file</a>";
    echo " || <a href='".$_SESSION["redirect_uri"]."?list_email2=true'>List Email</a>";
//    echo " || <a href='".$_SESSION["redirect_uri"]."?send_email=true'>Send Email</a>";
//    echo " || <a href='".$_SESSION["redirect_uri"]."?logout=true'>Logout</a><br/><br/>\n\n";
    echo "<br/><br/>\n\n";
}

if(isset($_GET["logout"])) {
    flush_token();
    echo "Logged out<br/>";
    echo "<a href='".$_SESSION["redirect_uri"]."'>Start new session</a>";
    die();
}
else if(isset($_GET["profile"])) {
    view_profile();
}
else if(isset($_GET["refresh_token"])) {
    refresh_token();
}
else if(isset($_GET["list_log"])) {
    list_log();
}

else if(isset($_GET["list_email2"])) {
    list_email2();
}

else if(isset($_GET["check_email"])) {
    check_email();
}
else if(isset($_GET["send_email"])) {
    send_email();
}
else if(token()) {
    echo "<pre>"; print_r(token()); echo "</pre>";
}
elseif (isset($_GET["code"])) {
    echo "<pre>";print_r($_GET);echo "</pre>";
    $token_request_data = array (
        "grant_type" => "authorization_code",
        "code" => $_GET["code"],
        "redirect_uri" => $_SESSION["redirect_uri"],
        "scope" => implode(" ", $_SESSION["scopes"]),
        "client_id" => $_SESSION["client_id"],
        "client_secret" => $_SESSION["client_secret"]
    );
    $body = http_build_query($token_request_data);
    $response = runCurl($_SESSION["authority"].$_SESSION["token_url"], $body);
    $response = json_decode($response);

    store_token($response);
    file_put_contents(file_active_user_id, get_user_id());
    file_put_contents(file_access_token, $response->access_token);
//    header("Location: " . $_SESSION["redirect_uri"]);
}
else {
    $accessUrl = $_SESSION["authority"].$_SESSION["auth_url"];
    echo "<a href='$accessUrl'>Login with Office 365</a>";
}

function send_email() {

    $_SESSION["scopes"] = array("offline_access", "openid", "https://outlook.office.com/mail.send");

    $_SESSION["api_url"] = "https://outlook.office.com/api/v2.0/me/sendmail";

//    $_SESSION["access_token"] = iI0MzlkNjYwZS04ODBkLTRiY2EtODQ5Yy01NDZmY2E5MzIzMTQiLCJhcHBpZGFjciI6IjEiLCJlbmZwb2xpZHMiOltdLCJmYW1pbHlfbmFtZSI6ImFwaSIsImdpdmVuX25hbWUiOiJ0ZXN0IiwiaXBhZGRyIjoiODEuOTIuMjAwLjIxOSIsIm5hbWUiOiJ0ZXN0IGFwaSIsIm9pZCI6Ijg0MTEyZWMwLTE4ZjAtNDg3Yi1hZDhlLWQ2MzUwYzc3ODQ2MSIsInB1aWQiOiIxMDAzMjAwMTc2OTVCMDBDIiwicmgiOiIwLkFTOEFicEdiRHZYT3hVS18wcVVXcjNXdUFnNW1uVU1OaU1wTGhKeFViOHFUSXhRdkFLby4iLCJzY3AiOiJNYWlsLlJlYWQgTWFpbC5TZW5kIFVzZXIuUmVhZCIsInNpZCI6IjE0ODg1YzcwLWYwOTUtNGY0Ni1hNmUzLTZlNGJkODBmOWJkMCIsInN1YiI6IjV0QkFxbjBmTE45aE9Db3ZwRXJmWEZDNlY4elNuTHQwX2VqdmZtZkxkWVEiLCJ0aWQiOiIwZTliOTE2ZS1jZWY1LTQyYzUtYmZkMi1hNTE2YWY3NWFlMDIiLCJ1bmlxdWVfbmFtZSI6InRlc3RhcGlAdGVhbW9zbWluZS5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJ0ZXN0YXBpQHRlYW1vc21pbmUub25taWNyb3NvZnQuY29tIiwidXRpIjoibjljM213V3RIMEdDbHBXX21uUk9BQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il19.PzxI5urUdze-CZT4rnXsoK9ve4tnbFQ0Ky7c1ZqwUJ66t6FtPYFi6nlYt2buSnnd1VXHxekmzoZQPb40nnwJGbfsVvOutd_nouo3WyqFphnq9IoqUmoTSpvowW1JC1uVNf-Kv--_MRwvcxDs5ZzYt-l3ZGcl3zwWD97mv-XxO_QOk83NfpJZf3HGV3cHYAYlnya66P53xqG_wTl9A8iOX9RmMSRVGZxG6IiDbqkKiGejQTDCoJekBZhxnx5yjxqYWoM0pkMcXyHakFN2aetv-XPn6NquM6Wjg8r11nmPo4TB0J4ip7301w58LSO0620WSjEOVh2p6xUaru8VzCJJbw";
    $_SESSION["access_token"] = file_get_contents(file_access_token);


if(isset($_POST["send"])) {
    $to = array();
    $toFromForm = explode(";", $_POST["to"]);
    foreach ($toFromForm as $eachTo) {
        if(strlen(trim($eachTo)) > 0) {
            $thisTo = array(
                "EmailAddress" => array(
                    "Address" => trim($eachTo)
                )
            );
            array_push($to, $thisTo);
        }
    }
    if (count($to) == 0) {
        die("Need email address to send email");
    }

    $request = array(
        "Message" => array(
            "Subject" =>$_POST["subject"],
            "ToRecipients" => $to,
//            "Attachments" => $attachments,
            "Body" => array(
                "ContentType" => "HTML",
                "Content" => utf8_encode($_POST["message"])
            )
        )
    );

    $request = json_encode($request);
    $headers = array(
        "User-Agent: php-tutorial/1.0",
        "Authorization: Bearer ".$_SESSION["access_token"],
        "Accept: application/json",
        "Content-Type: application/json",
        "Content-Length: ". strlen($request)
    );

    $response = runCurl($_SESSION["api_url"], $request, $headers);
    echo "<pre>"; print_r($response); echo "</pre>";
    /* if $response["code"] == 202 then mail sent successfully */
}

else {
?>
<form method="post" enctype="multipart/form-data" accept-charset="ISO-8859-1">
<table style="width: 1000px;">
    <tr>
        <td style="width: 150px;">To</td>
        <td style="width: 850px;"><input type="text" name="to" required value="" style="width: 100%;"/></td>
    </tr>
    <tr>
        <td>Subject</td>
        <td><input type="text" name="subject" required value="Prueba envio <?php echo date("d/m/Y H:i:s"); ?>" style="width: 100%;"/></td>
    </tr>
    <tr>
        <td style="vertical-align: top;">Message</td>
        <td><textarea name="message" required style="width: 100%; height: 60px;"></textarea></td>
    </tr>
    <tr>
        <td></td>
        <td><input type="submit" name="send" value="Enviar"/></td>
    </tr>
</table>
</form>
<?php
}
}

function check_email() {
    $token_request_data = array (
        "grant_type" => "refresh_token",
        "refresh_token" => token()->refresh_token,
        "redirect_uri" => $_SESSION["redirect_uri"],
        "scope" => implode(" ", $_SESSION["scopes"]),
        "client_id" => $_SESSION["client_id"],
        "client_secret" => $_SESSION["client_secret"]
    );
    $body = http_build_query($token_request_data);
    $response = runCurl($_SESSION["authority"].$_SESSION["token_url"], $body);
    $response = json_decode($response);
    store_token($response);
    file_put_contents(file_access_token, $response->access_token);

	view_profile2();

    $headers = array(
        "User-Agent: php-tutorial/1.0",
        "Authorization: Bearer ".token()->access_token,
        "Accept: application/json",
        "client-request-id: ".makeGuid(),
        "return-client-request-id: true",
        "X-AnchorMailbox: ". get_user_email()
    );
    $top = 20000;
    $skip = isset($_GET["skip"]) ? intval($_GET["skip"]) : 0;
    $search = array (
        // Only return selected fields
        "\$select" => "Subject,ReceivedDateTime,Sender,From,ToRecipients,HasAttachments,BodyPreview",
        // Sort by ReceivedDateTime, newest first
        "\$orderby" => "ReceivedDateTime DESC",
        // Return at most n results
        "\$top" => $top, "\$skip" => $skip
    );
    $outlookApiUrl = $_SESSION["api_url"] . "/Me/MailFolders/Inbox/Messages?" . http_build_query($search);
    $response = runCurl($outlookApiUrl, null, $headers);
    $response = explode("\n", trim($response));
    $response = $response[count($response) - 1];
    $response = json_decode($response, true);
    //echo "<pre>"; print_r($response); echo "</pre>";
	$now = new DateTime(null, new DateTimeZone('Europe/Madrid'));
    if(isset($response["value"]) && count($response["value"]) > 0) {
 //       echo "<style type='text/css'>td{border: 2px solid #cccccc;padding: 30px;text-align: center;vertical-align: top;}</style>";
 //       echo "<table style='width: 100%;'><tr><th>From</th><th>Subject</th><th>Preview</th></tr>";
        $numVisitors = 0;
        foreach ($response["value"] as $mail) {
        $numVisitors=$numVisitors+1;
          	}
            echo $now->format('Y-m-d H:i:s');
            echo "<div><h3><i>You have $numVisitors Messages</i></h3></div>";
            file_put_contents(file_user_log, "\n" . $now->format('Y-m-d H:i:s') . " $numVisitors" . PHP_EOL, FILE_APPEND);
 /*           $BodyPreview = str_replace("\n", "<br/>", $mail["BodyPreview"]);
            echo "<tr>";
            echo "<td>".$mail["From"]["EmailAddress"]["Address"].
                "<br/><a target='_blank' href='?view_email=".$mail["Id"]."'>View Email</a>";
            if($mail["HasAttachments"] == 1) {
                echo "<br/><a target='_blank' href='?view_attachments=".$mail["Id"]."'>View Attachments</a>";
            }
            echo "</td><td>".$mail["Subject"]."</td>";
            echo "<td>".$BodyPreview."</td>";
            echo "</tr>";
        }
        echo "</table>";
    } */
    }
    else {
        echo "<div><h3><i>No email found</i></h3></div>";
        file_put_contents(file_user_log, "\n" . $now->format('Y-m-d H:i:s') . " 0" . PHP_EOL, FILE_APPEND);
    }
   /* $prevLink = "";
    if($skip > 0) {
        $prev = $skip - $top;
        $prevLink = "<a href='?list_email=true&skip=".$prev."'>Previous Page</a>";
    }
    if(isset($response["@odata.nextLink"])) {
        if($prevLink != "") {
            $prevLink .= " ||| ";
        }
        echo "<br/>".$prevLink."<a href='?list_email2=true&skip=".($skip + $top)."'>Next Page</a>";
    }
    else {
        echo "<br/>" . $prevLink;
    } */
    die();
}


function list_email2() {
    $headers = array(
        "User-Agent: php-tutorial/1.0",
        "Authorization: Bearer ".token()->access_token,
        "Accept: application/json",
        "client-request-id: ".makeGuid(),
        "return-client-request-id: true",
        "X-AnchorMailbox: ". get_user_email()
    );
    $top = 20000;
    $skip = isset($_GET["skip"]) ? intval($_GET["skip"]) : 0;
    $search = array (
        // Only return selected fields
        "\$select" => "Subject,ReceivedDateTime,Sender,From,ToRecipients,HasAttachments,BodyPreview",
        // Sort by ReceivedDateTime, newest first
        "\$orderby" => "ReceivedDateTime DESC",
        // Return at most n results
        "\$top" => $top, "\$skip" => $skip
    );
    $outlookApiUrl = $_SESSION["api_url"] . "/Me/MailFolders/Inbox/Messages?" . http_build_query($search);
    $response = runCurl($outlookApiUrl, null, $headers);
    $response = explode("\n", trim($response));
    $response = $response[count($response) - 1];
    $response = json_decode($response, true);
    //echo "<pre>"; print_r($response); echo "</pre>";
	$now = new DateTime(null, new DateTimeZone('Europe/Madrid'));
    if(isset($response["value"]) && count($response["value"]) > 0) {
 //       echo "<style type='text/css'>td{border: 2px solid #cccccc;padding: 30px;text-align: center;vertical-align: top;}</style>";
 //       echo "<table style='width: 100%;'><tr><th>From</th><th>Subject</th><th>Preview</th></tr>";
        $numVisitors = 0;
        foreach ($response["value"] as $mail) {
        $numVisitors=$numVisitors+1;
          	}
            echo $now->format('Y-m-d H:i:s');
            echo "<div><h3><i>You have $numVisitors Messages</i></h3></div>";
            file_put_contents(file_user_log, "\n" . $now->format('Y-m-d H:i:s') . " $numVisitors" . PHP_EOL, FILE_APPEND);
 /*           $BodyPreview = str_replace("\n", "<br/>", $mail["BodyPreview"]);
            echo "<tr>";
            echo "<td>".$mail["From"]["EmailAddress"]["Address"].
                "<br/><a target='_blank' href='?view_email=".$mail["Id"]."'>View Email</a>";
            if($mail["HasAttachments"] == 1) {
                echo "<br/><a target='_blank' href='?view_attachments=".$mail["Id"]."'>View Attachments</a>";
            }
            echo "</td><td>".$mail["Subject"]."</td>";
            echo "<td>".$BodyPreview."</td>";
            echo "</tr>";
        }
        echo "</table>";
    } */
    }
    else {
        echo "<div><h3><i>No email found</i></h3></div>";
        file_put_contents(file_user_log, "\n" . $now->format('Y-m-d H:i:s') . " 0" . PHP_EOL, FILE_APPEND);
    }
   /* $prevLink = "";
    if($skip > 0) {
        $prev = $skip - $top;
        $prevLink = "<a href='?list_email=true&skip=".$prev."'>Previous Page</a>";
    }
    if(isset($response["@odata.nextLink"])) {
        if($prevLink != "") {
            $prevLink .= " ||| ";
        }
        echo "<br/>".$prevLink."<a href='?list_email2=true&skip=".($skip + $top)."'>Next Page</a>";
    }
    else {
        echo "<br/>" . $prevLink;
    } */
}



function list_log() {
//	header("Content-Type: text/plain");
//    readfile("officemanvgsm_log.txt");
// echo file_get_contents("officemanvgsm_log.txt");
  $str = file_get_contents(file_user_log);
  echo preg_replace('!\r?\n!', '<br>', $str);
// refreshPage();
// setInterval("refreshPage()", 60000);
}

function refresh_token() {
    $token_request_data = array (
        "grant_type" => "refresh_token",
        "refresh_token" => token()->refresh_token,
        "redirect_uri" => $_SESSION["redirect_uri"],
        "scope" => implode(" ", $_SESSION["scopes"]),
        "client_id" => $_SESSION["client_id"],
        "client_secret" => $_SESSION["client_secret"]
    );
    $body = http_build_query($token_request_data);
    $response = runCurl($_SESSION["authority"].$_SESSION["token_url"], $body);
    $response = json_decode($response);
    store_token($response);
    file_put_contents(file_access_token, $response->access_token);
//    header("Location: " . $_SESSION["redirect_uri"]);
}

function get_user_id() {
    if(isset($_SESSION["user_id"]) && strlen($_SESSION["user_id"]) > 0) {
        return $_SESSION["user_id"];
    }
    view_profile(true);
    $response = json_decode(file_get_contents(file_user_data));
    $_SESSION["user_id"] = $response->Id;
    return $response->Id;
}

function get_user_email() {
    if(isset($_SESSION["user_email"]) && strlen($_SESSION["user_email"]) > 0) {
        return $_SESSION["user_email"];
    }
    view_profile(true);
    $response = json_decode(file_get_contents(file_user_data));
    $_SESSION["user_email"] = $response->EmailAddress;
    return $response->EmailAddress;
}

function view_profile2($skipPrint = false) {
    $headers = array(
        "User-Agent: php-tutorial/1.0",
        "Authorization: Bearer ".token()->access_token,
        "Accept: application/json",
        "client-request-id: ".makeGuid(),
        "return-client-request-id: true"
    );
    $outlookApiUrl = $_SESSION["api_url"] . "/Me";
    $response = runCurl($outlookApiUrl, null, $headers);
    $response = explode("\n", trim($response));
    $response = $response[count($response) - 1];
    file_put_contents(file_user_data, $response);
    $response = json_decode($response);
    $_SESSION["user_id"] = $response->Id;
    $_SESSION["mail_id"] = $response->MailboxGuid;
    $_SESSION["user_email"] = $response->EmailAddress;
    if(!$skipPrint) {
//        echo "<pre>"; print_r($response); echo "</pre>";
    }
}

function view_profile($skipPrint = false) {
    $headers = array(
        "User-Agent: php-tutorial/1.0",
        "Authorization: Bearer ".token()->access_token,
        "Accept: application/json",
        "client-request-id: ".makeGuid(),
        "return-client-request-id: true"
    );
    $outlookApiUrl = $_SESSION["api_url"] . "/Me";
    $response = runCurl($outlookApiUrl, null, $headers);
    $response = explode("\n", trim($response));
    $response = $response[count($response) - 1];
    file_put_contents(file_user_data, $response);
    $response = json_decode($response);
    $_SESSION["user_id"] = $response->Id;
    $_SESSION["mail_id"] = $response->MailboxGuid;
    $_SESSION["user_email"] = $response->EmailAddress;
    if(!$skipPrint) {
        echo "<pre>"; print_r($response); echo "</pre>";
    }
}

function makeGuid(){
    if (function_exists('com_create_guid')) {
        error_log("Using 'com_create_guid'.");
        return strtolower(trim(com_create_guid(), '{}'));
    }
    else {
        $charid = strtolower(md5(uniqid(rand(), true)));
        $hyphen = chr(45);
        $uuid = substr($charid, 0, 8).$hyphen
            .substr($charid, 8, 4).$hyphen
            .substr($charid, 12, 4).$hyphen
            .substr($charid, 16, 4).$hyphen
            .substr($charid, 20, 12);
        return $uuid;
    }
}

function flush_token() {
    file_put_contents(file_auth_config, "");
    $_SESSION["user_id"] = "";
    $_SESSION["mail_id"] = "";
}

function store_token($o) {
    file_put_contents(file_auth_config, json_encode($o));
}

function token() {
    $text = file_exists(file_auth_config) ? file_get_contents(file_auth_config) : null;
    if($text != null && strlen($text) > 0) {
        return json_decode($text);
    }
    return null;
}

function init() {
    $_SESSION["client_id"] = CLIENTE_ID;
    $_SESSION["client_secret"] = SECRETO_ID;
    $_SESSION["redirect_uri"] = CLIENTE_URL;
    $_SESSION["authority"] = "https://login.microsoftonline.com";
    $_SESSION["scopes"] = array("offline_access", "openid");
    /* If you need to read email, then need to add following scope */
    if(true) {
        array_push($_SESSION["scopes"], "https://outlook.office.com/mail.read");
    }
    /* If you need to send email, then need to add following scope */
    if(true) {
        array_push($_SESSION["scopes"], "https://outlook.office.com/mail.send");
    }

    $_SESSION["auth_url"] = "/common/oauth2/v2.0/authorize";
    $_SESSION["auth_url"] .= "?client_id=".$_SESSION["client_id"];
    $_SESSION["auth_url"] .= "&redirect_uri=".$_SESSION["redirect_uri"];
    $_SESSION["auth_url"] .= "&response_type=code&scope=".implode(" ", $_SESSION["scopes"]);

    $_SESSION["token_url"] = "/common/oauth2/v2.0/token";

    $_SESSION["api_url"] = "https://outlook.office.com/api/v2.0";
}

function runCurl($url, $post = null, $headers = null) {
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_POST, $post == null ? 0 : 1);
    if($post != null) {
        curl_setopt($ch, CURLOPT_POSTFIELDS, $post);
    }
    curl_setopt($ch, CURLOPT_HTTPAUTH, CURLAUTH_ANY);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    if($headers != null) {
        curl_setopt($ch, CURLOPT_HEADER, true);
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    }
    $response = curl_exec($ch);
    $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    if($http_code >= 400) {
        echo "Error executing request to Office365 api with error code=$http_code<br/><br/>\n\n";
        echo "<pre>"; print_r($response); echo "</pre>";
        die();
    }
    return $response;
}
?>
