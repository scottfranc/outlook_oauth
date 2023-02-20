<?php
// Prerequisites:
// The Client ID, The Tenant ID, The secret client.

$TENANT_ID="b55...028";
$CLIENT_ID="72f...978";
$SCOPE="https://outlook.office365.com/IMAP.AccessAsUser.All"; //This is specifically for IMAP connection, you will need to replace this if you are using POP3 or SMTP.
$REDIRECT_URL="https://website.tld/getAtokenNet"; //you can change this to be anything, preferably on your domain

$authUri = 'https://login.microsoftonline.com/' . $TENANT_ID
           . '/oauth2/v2.0/authorize?client_id=' . $CLIENT_ID
           . '&scope=' . $SCOPE
           . '&REDIRECT_URL=' . urlencode($REDIRECT_URL)
           . '&response_type=code'
           . '&approval_prompt=auto';

echo($authUri); //this would print a URL string for you to copy into the browser and visit, when you visit this link, you will need to log into your Microsoft Exchange account or be already logged it into it.

//Once it is done, you should have a link in the address bar like "https://website.tld/getAtokenNet?code=ioPFnco...&session_state=d80fw1..." A code (remove the "&" at the end) and a session_state ID

//Next step is to fetch an access token


$CLIENT_SECRET="Y~tN...";
$SCOPE="https://outlook.office365.com/IMAP.AccessAsUser.All offline_access"; //NOTE the "offline_access" keyword.
$CODE="ioPFnco...";
$SESSION="d80fw1...";

echo "Authenticating session...";

$url= "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token";

$load_refresh_and_access_token = [ 
 'client_id'=>$CLIENT_ID,
 'scope'=>$SCOPE,
 'code'=>$CODE,
 'session_state'=>$SESSION,
 'client_secret'=>$CLIENT_SECRET,
 'REDIRECT_URL'=>$REDIRECT_URL,
 'grant_type'=>'authorization_code' ];

$ch=curl_init();
curl_setopt($ch,CURLOPT_URL,$url);
curl_setopt($ch,CURLOPT_POSTFIELDS, http_build_query($load_refresh_and_access_token));
curl_setopt($ch,CURLOPT_POST, 1);
curl_setopt($ch,CURLOPT_RETURNTRANSFER, true);

$tokenResult=curl_exec($ch);

echo "access token and refresh code : \n";

var_dump($tokenResult); //here you will have the access token and refresh code. 

//we then connect to the Microsoft Outlook Server using IMAP Protocol like below:

$mailbox = '{outlook.office365.com:993/imap/ssl}';
$username = 'email@company.tld'; //your email address
$accessToken = "eyh(S...FA";

$inbox = imap2_open($mailbox, $username, $accessToken, OP_XOAUTH2);

imap2_reopen($inbox, $mailbox.'INBOX');

$emails = imap2_search($inbox, 'ALL');


print_r($emails);

//With the access_code format above, we will need to always request a new token when it expires, so to avoid that can use the "refresh code" we generated earlier like below:


    $CLIENT_SECRET="i5H8Q...";
    $REFRESH_TOKEN="gedv34grv...";
    
    $url= "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token";
    
    $load_refresh_and_access_token = [ 
     'client_id'=>$CLIENT_ID,
     'client_secret'=>$CLIENT_SECRET,
     'refresh_token'=>$REFRESH_TOKEN,
     'grant_type'=>'refresh_token' ];
    
    $ch=curl_init();
    
    curl_setopt($ch,CURLOPT_URL,$url);
    curl_setopt($ch,CURLOPT_POSTFIELDS, http_build_query($load_refresh_and_access_token));
    curl_setopt($ch,CURLOPT_POST, 1);
    curl_setopt($ch,CURLOPT_RETURNTRANSFER, true);
    
    $tokenResult=curl_exec($ch);
    
    echo("Trying to get the token.... \n");
    
    if(!empty($tokenResult)){
        $decoded_result = json_decode($tokenResult,true);
        $accessToken = $decoded_result["accessToken"];
        if( !empty($accessToken) ){
            $inbox = imap2_open($mailbox, $username, $accessToken, OP_XOAUTH2); //here we fetch the emails, and can be fetched anytime without needing to renew the access token manually
        }
    }
