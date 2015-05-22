<?php

#from php-form-builder-class

class Spreadsheet {

    private $token;
    private $spreadsheet;
    private $worksheet;
    private $spreadsheetid;
    private $worksheetid;

    public function __construct($username, $password) {
        $this->authenticate($username, $password);
    }

    public function authenticate($username, $password) {
        $url = "https://www.google.com/accounts/ClientLogin";
        $fields = array(
            "accountType" => "HOSTED_OR_GOOGLE",
            "Email" => $username,
            "Passwd" => $password,
            "service" => "wise",
            "source" => "pfbc"
        );
        $curl = curl_init();
        curl_setopt($curl, CURLOPT_URL, $url);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($curl, CURLOPT_POST, true);
        curl_setopt($curl, CURLOPT_POSTFIELDS, $fields);
        $response = curl_exec($curl);
        $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
        curl_close($curl);

        if ($status == 200) {
            if (stripos($response, "auth=") !== false) {
                preg_match("/auth=([a-z0-9_\-]+)/i", $response, $matches);
                $this->token = $matches[1];
            }
        }
    }

    public function setSpreadsheet($title) {
        $this->spreadsheet = $title;
        return $this;
    }

    public function setSpreadsheetId($id) {
        $this->spreadsheetid = $id;
        return $this;
    }

    public function setWorksheet($title) {
        $this->worksheet = $title;
        return $this;
    }

    public function add($data) {
        if (!empty($this->token)) {
            $url = $this->getPostUrl();
            if (!empty($url)) {
                $headers = array(
                    "Content-Type: application/atom+xml",
                    "Authorization: GoogleLogin auth=" . $this->token,
                    "GData-Version: 3.0"
                );

                $columnIDs = $this->getColumnIDs();

                if ($columnIDs) {
                    $fields = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">';
                    foreach ($data as $key => $value) {
                        $key = $this->formatColumnID($key);
                        if (in_array($key, $columnIDs))
                            $fields .= "<gsx:$key><![CDATA[$value]]></gsx:$key>";
                    }
                    $fields .= '</entry>';

                    $curl = curl_init();
                    curl_setopt($curl, CURLOPT_URL, $url);
                    curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
                    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
                    curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
                    curl_setopt($curl, CURLOPT_POST, true);
                    curl_setopt($curl, CURLOPT_POSTFIELDS, $fields);
                    $response = curl_exec($curl);
                    $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
                    curl_close($curl);
                }
            }
        }
    }

    private function getColumnIDs() {
        $url = "https://spreadsheets.google.com/feeds/cells/" . $this->spreadsheetid . "/" . $this->worksheetid . "/private/full?max-row=1";
        $headers = array(
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        $curl = curl_init();
        curl_setopt($curl, CURLOPT_URL, $url);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
        $response = curl_exec($curl);

        $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
        curl_close($curl);

        if ($status == 200) {

            $columnIDs = array();
            $xml = simplexml_load_string($response);
            if ($xml->entry) {
                $columnSize = sizeof($xml->entry);
                for ($c = 0; $c < $columnSize; ++$c)
                    $columnIDs[] = $this->formatColumnID($xml->entry[$c]->content);
            }
            return $columnIDs;
        }

        return "";
    }

    private function getPostUrl() {
        if (empty($this->spreadsheetid)) {

            #find the id based on the spreadsheet name

            $url = "https://spreadsheets.google.com/feeds/spreadsheets/private/full?title=" . urlencode($this->spreadsheet);
            $headers = array(
                "Authorization: GoogleLogin auth=" . $this->token,
                "GData-Version: 3.0"
            );
            $curl = curl_init();
            curl_setopt($curl, CURLOPT_URL, $url);
            curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
            curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
            curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
            $response = curl_exec($curl);
            $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);

            if ($status == 200) {
                $spreadsheetXml = simplexml_load_string($response);
                if ($spreadsheetXml->entry) {
                    $this->spreadsheetid = basename(trim($spreadsheetXml->entry[0]->id));
                    $url = "https://spreadsheets.google.com/feeds/worksheets/" . $this->spreadsheetid . "/private/full";
                    if (!empty($this->worksheet))
                        $url .= "?title=" . $this->worksheet;

                    curl_setopt($curl, CURLOPT_URL, $url);
                    $response = curl_exec($curl);
                    $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
                    if ($status == 200) {
                        $worksheetXml = simplexml_load_string($response);
                        if ($worksheetXml->entry)
                            $this->worksheetid = basename(trim($worksheetXml->entry[0]->id));
                    }
                }
            }
            curl_close($curl);
        }


        if (!empty($this->spreadsheetid) && !empty($this->worksheetid))
            return "https://spreadsheets.google.com/feeds/list/" . $this->spreadsheetid . "/" . $this->worksheetid . "/private/full";

        return "";
    }

    private function formatColumnID($val) {
        return preg_replace("/[^a-zA-Z0-9.-]/", "", strtolower($val));
    }

    public function read() {
        if (!empty($this->token)) {
            $url = $this->getPostUrl();
            if (!empty($url)) {
                $report = array();
                //$url = "https://spreadsheets.google.com/feeds/cells/" . $this->spreadsheetid . "/" . $this->worksheetid . "/private/full?&alt=json"; // for old Google Sheets
                $url = "https://spreadsheets.google.com/feeds/cells/" . $this->spreadsheetid . "/" . $this->worksheetid . "/public/values?alt=json"; // compatible with new google sheets, you must publish the sheet first
                $headers = array(
                    "Authorization: GoogleLogin auth=" . $this->token,
                    "GData-Version: 3.0"
                );
                $curl = curl_init();
                curl_setopt($curl, CURLOPT_URL, $url);
                curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
                curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
                curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
                $response = curl_exec($curl);

                $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
                curl_close($curl);

                if ($status == 200) {
                    $responseArray = json_decode($response, true);
                    $report = array_fill(0, (int) $responseArray['feed']['gs$colCount']['$t'], array_fill(0, (int) $responseArray['feed']['entry'][count($responseArray['feed']['entry']) - 1]['gs$cell']['row'], ''));
                    for ($i = 0; $i < count($responseArray['feed']['entry']); $i++) {
                        $columnNumber = ((int) $responseArray['feed']['entry'][$i]['gs$cell']['col']) - 1;
                        $rowNumber = ((int) $responseArray['feed']['entry'][$i]['gs$cell']['row']) - 1;
                        if ($rowNumber === 0) {
                            $headers[$columnNumber] = $responseArray['feed']['entry'][$i]['gs$cell']['$t'];
                        }
                        if (!isset($report[$headers[$columnNumber]])) {
                            $report[$headers[$columnNumber]] = array_fill(0, (int) $responseArray['feed']['entry'][count($responseArray['feed']['entry']) - 1]['gs$cell']['row'], '');
                        }
                        $report[$headers[$columnNumber]][$rowNumber] = $responseArray['feed']['entry'][$i]['gs$cell']['$t']; // associative named index
                        $report[$columnNumber][$rowNumber] = $responseArray['feed']['entry'][$i]['gs$cell']['$t']; // flat numeric index
                        $report[$responseArray['feed']['entry'][$i]['title']['$t']] = $responseArray['feed']['entry'][$i]['gs$cell']['$t']; // associative table index
                    }
                    ksort($report);
                    return $report;
                }
            }
        }
        return "";
    }

}

?>
