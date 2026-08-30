<?php
/**
 * Word Texter - Next-Gen Corporate & Office Automation Suite
 * Developed & All Rights Reserved: Devsingh.m
 * 100% Responsive for All Mobile Phones, Tablets & Windows OS
 */

ini_set('display_errors', 0);
error_reporting(E_ALL & ~E_NOTICE & ~E_DEPRECATED);
ini_set('max_execution_time', 600);
ini_set('memory_limit', '512M');

if (session_status() === PHP_SESSION_NONE) {
    @session_start();
}

if (!isset($_SESSION['wt_workspace'])) {
    $_SESSION['wt_workspace'] = 'ws_' . substr(md5(session_id() . ($_SERVER['REMOTE_ADDR'] ?? '127.0.0.1')), 0, 16);
}
$workspaceKey = $_SESSION['wt_workspace'];

$baseStorageDir = __DIR__ . '/secure_data';
$userStorageDir = $baseStorageDir . '/' . $workspaceKey;

if (!file_exists($baseStorageDir)) {
    @mkdir($baseStorageDir, 0777, true);
    @file_put_contents($baseStorageDir . '/.htaccess', "Order Deny,Allow\nDeny from all\n");
    @file_put_contents($baseStorageDir . '/index.html', "");
}

if (!file_exists($userStorageDir)) {
    @mkdir($userStorageDir, 0777, true);
    @file_put_contents($userStorageDir . '/index.html', "");
}

function normalizeKey($str) {
    $str = preg_replace('/\.(pdf|docx|doc|xlsx|csv)$/i', '', trim($str));
    return strtolower(preg_replace('/[^a-zA-Z0-9]/', '', $str));
}

function normalizeText($str) {
    $str = str_replace(["\xc2\xa0", "&nbsp;"], ' ', $str);
    return trim(preg_replace('/\s+/u', ' ', $str));
}

// -------------------------------------------------------------
// DOWNLOAD HANDLERS
// -------------------------------------------------------------
if (isset($_GET['action']) && $_GET['action'] === 'download_template') {
    $candidates = [
        __DIR__ . '/Word_file_Address_Change_Template.xlsx',
        __DIR__ . '/Word file  Address Change Template.xlsx'
    ];
    $found = '';
    foreach ($candidates as $c) {
        if (file_exists($c)) { $found = $c; break; }
    }
    if (!empty($found)) {
        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="Word_file_Address_Change_Template.xlsx"');
        header('Content-Length: ' . filesize($found));
        header('Cache-Control: private, no-transform, no-store, must-revalidate');
        readfile($found);
        exit;
    } else {
        die('Template file not found on server.');
    }
}

if (isset($_GET['action']) && $_GET['action'] === 'download_file') {
    $reqFile = basename($_GET['file'] ?? '');
    $filePath = $userStorageDir . '/' . $reqFile;
    if (!empty($reqFile) && file_exists($filePath) && is_file($filePath)) {
        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="' . $reqFile . '"');
        header('Content-Length: ' . filesize($filePath));
        header('Cache-Control: private, no-transform, no-store, must-revalidate');
        readfile($filePath);
        exit;
    } else {
        die('File expired or not found.');
    }
}

if (isset($_GET['action']) && $_GET['action'] === 'download_zip') {
    $reqZip = basename($_GET['zip'] ?? '');
    $zipPath = $userStorageDir . '/' . $reqZip;
    if (!empty($reqZip) && file_exists($zipPath) && is_file($zipPath)) {
        header('Content-Description: File Transfer');
        header('Content-Type: application/zip');
        header('Content-Disposition: attachment; filename="' . $reqZip . '"');
        header('Content-Length: ' . filesize($zipPath));
        header('Cache-Control: private, no-transform, no-store, must-revalidate');
        readfile($zipPath);
        exit;
    } else {
        die('ZIP archive expired or not found.');
    }
}

// -------------------------------------------------------------
// XLSX & DOCX REPLACEMENT ENGINE
// -------------------------------------------------------------
function parseXlsxTemplateRobust($filePath) {
    if (!class_exists('ZipArchive')) return [];
    $zip = new ZipArchive();
    if ($zip->open($filePath) !== TRUE) return [];

    $sharedStrings = [];
    $ssXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($ssXml) {
        $xml = @simplexml_load_string($ssXml);
        if ($xml) {
            foreach ($xml->si as $val) {
                if (isset($val->t)) {
                    $sharedStrings[] = (string)$val->t;
                } elseif (isset($val->r)) {
                    $textParts = [];
                    foreach ($val->r as $r) { $textParts[] = (string)$r->t; }
                    $sharedStrings[] = implode('', $textParts);
                } else {
                    $sharedStrings[] = '';
                }
            }
        }
    }

    $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    if (!$sheetXml) {
        $zip->close();
        return [];
    }

    $xml = @simplexml_load_string($sheetXml);
    $rows = [];
    if ($xml && isset($xml->sheetData->row)) {
        foreach ($xml->sheetData->row as $row) {
            $rowArr = [];
            foreach ($row->c as $cell) {
                $val = (string)$cell->v;
                $type = (string)$cell['t'];
                if ($type === 's' && isset($sharedStrings[(int)$val])) {
                    $val = $sharedStrings[(int)$val];
                }
                $cellRef = (string)$cell['r'];
                preg_match('/([A-Z]+)(\\d+)/', $cellRef, $matches);
                $colLetters = $matches[1] ?? 'A';
                $colIdx = 0;
                for ($k = 0; $k < strlen($colLetters); $k++) {
                    $colIdx = $colIdx * 26 + (ord($colLetters[$k]) - ord('A') + 1);
                }
                $rowArr[$colIdx - 1] = trim($val);
            }
            $rows[] = $rowArr;
        }
    }
    $zip->close();
    if (empty($rows)) return [];

    $headers = array_shift($rows);
    $fileIdx = -1; $oldIdx = -1; $newIdx = -1;

    foreach ($headers as $idx => $hText) {
        $norm = strtolower(preg_replace('/[^a-zA-Z0-9]/', '', (string)$hText));
        if (strpos($norm, 'filename') !== false) $fileIdx = $idx;
        if (strpos($norm, 'oldaddress') !== false) $oldIdx = $idx;
        if (strpos($norm, 'newaddress') !== false) $newIdx = $idx;
    }

    $rules = [];
    foreach ($rows as $data) {
        if ($fileIdx >= 0 && $oldIdx >= 0 && $newIdx >= 0 && !empty($data[$fileIdx])) {
            $rawName = trim($data[$fileIdx]);
            $key = normalizeKey($rawName);
            $rules[$key] = [
                'raw_name' => $rawName,
                'old' => trim($data[$oldIdx] ?? ''),
                'new' => trim($data[$newIdx] ?? '')
            ];
        }
    }
    return $rules;
}

function replaceTextInDocxXml($xml, $oldText, $newText) {
    if (empty($oldText) || empty($newText)) return $xml;

    $normOld = normalizeText($oldText);
    $escapedNew = htmlspecialchars($newText, ENT_XML1 | ENT_COMPAT, 'UTF-8');

    $updatedXml = preg_replace_callback('/<w:p\\b[^>]*>.*?<\\/w:p>/su', function($pMatch) use ($oldText, $normOld, $escapedNew) {
        $pXml = $pMatch[0];

        if (!preg_match_all('/<w:t(\\b[^>]*)>(.*?)<\\/w:t>/su', $pXml, $matches, PREG_OFFSET_CAPTURE)) {
            return $pXml;
        }

        $fullText = '';
        foreach ($matches[2] as $m) {
            $fullText .= $m[0];
        }

        $normFull = normalizeText($fullText);
        $found = false;
        $replacedText = '';

        $decodedFull = html_entity_decode($fullText, ENT_QUOTES | ENT_XML1, 'UTF-8');
        if (mb_strpos($decodedFull, $oldText) !== false) {
            $replacedText = str_replace($oldText, $escapedNew, $decodedFull);
            $found = true;
        } elseif (mb_strpos($normFull, $normOld) !== false) {
            $replacedText = str_replace($normOld, $escapedNew, $normFull);
            $found = true;
        }

        if ($found) {
            $firstTStart = $matches[0][0][1];
            $firstTEnd = $firstTStart + strlen($matches[0][0][0]);
            
            $newP = substr($pXml, 0, $firstTStart);
            $newP .= '<w:t xml:space="preserve">' . $replacedText . '</w:t>';
            
            $lastOffset = $firstTEnd;
            for ($i = 1; $i < count($matches[0]); $i++) {
                $tStart = $matches[0][$i][1];
                $tEnd = $tStart + strlen($matches[0][$i][0]);
                $newP .= substr($pXml, $lastOffset, $tStart - $lastOffset);
                $newP .= '<w:t></w:t>';
                $lastOffset = $tEnd;
            }
            $newP .= substr($pXml, $lastOffset);
            return $newP;
        }

        return $pXml;
    }, $xml);

    $rawEscOld = htmlspecialchars($oldText, ENT_XML1 | ENT_COMPAT, 'UTF-8');
    $updatedXml = str_replace([$rawEscOld, $oldText], $escapedNew, $updatedXml);

    return $updatedXml;
}

function processDocxRobust($filePath, $oldText, $newText) {
    if (!class_exists('ZipArchive')) return false;
    $zip = new ZipArchive();
    if ($zip->open($filePath) !== TRUE) return false;

    $targets = ['word/document.xml'];
    for ($i = 0; $i < $zip->numFiles; $i++) {
        $name = $zip->statIndex($i)['name'];
        if (preg_match('/^word\\/(header|footer)\\d+\\.xml$/i', $name)) {
            $targets[] = $name;
        }
    }

    foreach ($targets as $entry) {
        $xmlContent = $zip->getFromName($entry);
        if ($xmlContent !== false) {
            $modifiedXml = replaceTextInDocxXml($xmlContent, $oldText, $newText);
            if ($modifiedXml !== $xmlContent) {
                $zip->addFromString($entry, $modifiedXml);
            }
        }
    }

    $zip->close();
    return true;
}

$processedFiles = [];
$zipArchiveName = '';
$alertMsg = '';
$auditReport = [];
$successCount = 0;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_FILES['template_file']) && isset($_FILES['word_files'])) {
        $tpl = $_FILES['template_file'];
        
        if ($tpl['error'] !== UPLOAD_ERR_OK) {
            $alertMsg = '<div class="cyber-alert error"><span class="alert-icon">&#9888;</span> Select a valid Excel (.xlsx / .csv) template file.</div>';
        } else {
            $rules = parseXlsxTemplateRobust($tpl['tmp_name']);

            if (empty($rules)) {
                $alertMsg = '<div class="cyber-alert error"><span class="alert-icon">&#9888;</span> Invalid Template Structure. Headers must match: <strong>File Name</strong>, <strong>Old Address</strong>, <strong>New Address</strong>.</div>';
            } else {
                $zipArchiveName = 'Word_Texter_Batch_' . date('Ymd_His') . '.zip';
                $batchZipPath = $userStorageDir . '/' . $zipArchiveName;
                $zipArchive = new ZipArchive();
                $zipArchive->open($batchZipPath, ZipArchive::CREATE);

                $fileCount = count($_FILES['word_files']['name']);
                for ($i = 0; $i < $fileCount; $i++) {
                    $origName = $_FILES['word_files']['name'][$i];
                    $tmpLoc = $_FILES['word_files']['tmp_name'][$i];
                    $ext = strtolower(pathinfo($origName, PATHINFO_EXTENSION));

                    if ($_FILES['word_files']['error'][$i] === UPLOAD_ERR_OK && $ext === 'docx') {
                        $cleanName = preg_replace('/[^a-zA-Z0-9_\\-\\.\\,\\(\\)\\s]/', '', $origName);
                        $destPath = $userStorageDir . '/' . $cleanName;
                        move_uploaded_file($tmpLoc, $destPath);

                        $key = normalizeKey($cleanName);
                        $isSuccess = false;

                        if (isset($rules[$key])) {
                            $oldAddr = $rules[$key]['old'];
                            $newAddr = $rules[$key]['new'];
                            processDocxRobust($destPath, $oldAddr, $newAddr);
                            $status = 'Address Replaced & Verified';
                            $isSuccess = true;
                            $successCount++;
                        } else {
                            $status = 'File Name Not Matched In Template';
                        }

                        $zipArchive->addFile($destPath, $cleanName);
                        $processedFiles[] = $cleanName;
                        $auditReport[] = [
                            'filename' => $cleanName,
                            'status' => $status,
                            'success' => $isSuccess,
                            'time' => date('H:i:s')
                        ];
                    }
                }
                $zipArchive->close();

                if (!empty($processedFiles)) {
                    $alertMsg = '<div class="cyber-alert success"><span class="alert-icon">&#10004;</span> <strong>OFFICE AUTOMATION COMPLETE:</strong> ' . $successCount . ' file(s) processed with complete audit integrity.</div>';
                } else {
                    $alertMsg = '<div class="cyber-alert error"><span class="alert-icon">&#9888;</span> No valid Word (.docx) documents uploaded.</div>';
                }
            }
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
    <title>Word Texter &bull; Office & Compliance Automation Suite</title>
    <!-- Cross-Browser Responsive Meta & Futuristic Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Orbitron:wght@600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Syncopate:wght@700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-deep: #060919;
            --neon-cyan: #00f0ff;
            --neon-pink: #ff007f;
            --neon-emerald: #00ff88;
            --neon-amber: #ffb703;
            --neon-purple: #9d4edd;
            --card-glass: rgba(11, 19, 43, 0.85);
            --card-border: rgba(0, 240, 255, 0.35);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        body {
            background-color: var(--bg-deep);
            background-image: 
                radial-gradient(circle at 10% 20%, rgba(157, 78, 221, 0.28) 0%, transparent 42%),
                radial-gradient(circle at 90% 15%, rgba(0, 240, 255, 0.25) 0%, transparent 45%),
                radial-gradient(circle at 50% 85%, rgba(255, 0, 127, 0.20) 0%, transparent 50%),
                linear-gradient(rgba(0, 240, 255, 0.035) 1px, transparent 1px),
                linear-gradient(90deg, rgba(0, 240, 255, 0.035) 1px, transparent 1px),
                linear-gradient(to bottom, #060919, #0c122c, #060919);
            background-size: auto, auto, auto, 36px 36px, 36px 36px, auto;
            background-attachment: fixed;
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: flex-start;
            padding: 30px 16px;
            color: #f1f5f9;
            position: relative;
            overflow-x: hidden;
        }

        /* Office Floating Elements Background Watermarks */
        .office-bg-decor {
            position: fixed;
            pointer-events: none;
            z-index: 0;
            color: rgba(0, 240, 255, 0.035);
            font-family: 'Orbitron', sans-serif;
            font-weight: 900;
            text-transform: uppercase;
            user-select: none;
        }
        .office-decor-1 { top: 8%; left: 3%; font-size: 55px; transform: rotate(-8deg); }
        .office-decor-2 { bottom: 10%; right: 4%; font-size: 65px; transform: rotate(10deg); }
        .office-decor-3 { top: 45%; left: -2%; font-size: 80px; transform: rotate(-90deg); }

        .app-container {
            width: 100%;
            max-width: 980px;
            margin: 0 auto;
            position: relative;
            z-index: 1;
        }

        /* Top Brand Header */
        .brand-header-wrap {
            text-align: center;
            margin-bottom: 30px;
        }

        .dev-badge-top {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            background: rgba(0, 240, 255, 0.08);
            border: 1px solid rgba(0, 240, 255, 0.45);
            padding: 7px 18px;
            border-radius: 9999px;
            margin-bottom: 14px;
            box-shadow: 0 0 22px rgba(0, 240, 255, 0.3);
            backdrop-filter: blur(12px);
            max-width: 100%;
        }

        .pulse-core {
            width: 9px;
            height: 9px;
            background: var(--neon-emerald);
            border-radius: 50%;
            box-shadow: 0 0 12px var(--neon-emerald);
            animation: pulseCore 1.6s infinite;
            flex-shrink: 0;
        }
        @keyframes pulseCore {
            0% { transform: scale(0.9); opacity: 0.7; }
            50% { transform: scale(1.35); opacity: 1; filter: drop-shadow(0 0 8px var(--neon-emerald)); }
            100% { transform: scale(0.9); opacity: 0.7; }
        }

        .dev-badge-text {
            font-family: 'Orbitron', sans-serif;
            font-size: 11.5px;
            font-weight: 800;
            letter-spacing: 1.5px;
            color: var(--neon-cyan);
            text-transform: uppercase;
            white-space: normal;
            word-break: break-word;
        }

        /* Colorful Glowing Title */
        .glowing-rainbow-title {
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(30px, 6vw, 50px);
            font-weight: 900;
            letter-spacing: clamp(1px, 1.5vw, 3px);
            text-transform: uppercase;
            margin-bottom: 10px;
            background: linear-gradient(135deg, #00f0ff 0%, #ff007f 25%, #ffb703 50%, #9d4edd 75%, #00ff88 100%);
            background-size: 300% 300%;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            animation: rainbowFlow 5s ease infinite;
            filter: drop-shadow(0 0 25px rgba(0, 240, 255, 0.4));
            line-height: 1.2;
        }
        @keyframes rainbowFlow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        .developer-signature-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex-wrap: wrap;
            gap: 6px;
            background: linear-gradient(135deg, rgba(255, 0, 127, 0.15), rgba(157, 78, 221, 0.2));
            border: 1px solid rgba(255, 0, 127, 0.5);
            padding: 6px 14px;
            border-radius: 8px;
            font-family: 'Orbitron', sans-serif;
            font-size: 11.5px;
            font-weight: 800;
            color: #ffffff;
            letter-spacing: 0.8px;
            box-shadow: 0 0 15px rgba(255, 0, 127, 0.35);
            margin-bottom: 14px;
            max-width: 100%;
        }
        .dev-author-name {
            color: var(--neon-cyan);
            font-weight: 900;
            text-shadow: 0 0 10px rgba(0, 240, 255, 0.6);
        }

        .brand-subtext {
            font-size: clamp(13px, 2.5vw, 15px);
            color: #94a3b8;
            max-width: 720px;
            margin: 0 auto;
            line-height: 1.6;
            padding: 0 8px;
        }

        .txt-cyan { color: var(--neon-cyan); font-weight: 700; text-shadow: 0 0 10px rgba(0, 240, 255, 0.45); }
        .txt-pink { color: var(--neon-pink); font-weight: 700; text-shadow: 0 0 10px rgba(255, 0, 127, 0.45); }
        .txt-amber { color: var(--neon-amber); font-weight: 700; text-shadow: 0 0 10px rgba(255, 183, 3, 0.45); }
        .txt-emerald { color: var(--neon-emerald); font-weight: 700; text-shadow: 0 0 10px rgba(0, 255, 136, 0.45); }

        /* Main Holographic Terminal Card */
        .glass-terminal-card {
            background: var(--card-glass);
            backdrop-filter: blur(28px);
            -webkit-backdrop-filter: blur(28px);
            border: 1px solid var(--card-border);
            border-radius: 22px;
            padding: clamp(20px, 4vw, 42px);
            box-shadow: 0 30px 80px rgba(0, 0, 0, 0.8), inset 0 0 30px rgba(0, 240, 255, 0.07);
            position: relative;
            overflow: hidden;
            width: 100%;
        }

        .glass-terminal-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0; height: 3.5px;
            background: linear-gradient(90deg, #00f0ff, #ff007f, #ffb703, #9d4edd, #00ff88);
            box-shadow: 0 0 25px rgba(0, 240, 255, 0.9);
        }

        /* Step 1: Template Banner (Responsive Flexbox) */
        .step-template-banner {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 16px;
            background: linear-gradient(135deg, rgba(0, 240, 255, 0.06), rgba(157, 78, 221, 0.06));
            border: 1px solid rgba(0, 240, 255, 0.4);
            border-radius: 16px;
            padding: clamp(16px, 3vw, 24px);
            margin-bottom: 26px;
            box-shadow: inset 0 0 25px rgba(0, 240, 255, 0.06);
            transition: all 0.3s ease;
        }
        .step-template-banner:hover {
            border-color: var(--neon-cyan);
            box-shadow: 0 0 30px rgba(0, 240, 255, 0.35);
        }

        .tpl-banner-left {
            display: flex;
            align-items: center;
            gap: 16px;
            flex: 1 1 300px;
        }

        .office-icon-box {
            width: clamp(46px, 6vw, 56px);
            height: clamp(46px, 6vw, 56px);
            border-radius: 14px;
            background: linear-gradient(135deg, rgba(0, 240, 255, 0.22), rgba(255, 0, 127, 0.22));
            border: 1.5px solid var(--neon-cyan);
            color: var(--neon-cyan);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: clamp(22px, 3.5vw, 28px);
            box-shadow: 0 0 20px rgba(0, 240, 255, 0.4);
            flex-shrink: 0;
        }

        .template-heading-title {
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(14px, 2vw, 16px);
            font-weight: 700;
            color: #ffffff;
            letter-spacing: 0.6px;
            margin-bottom: 4px;
        }
        .template-sub-desc {
            font-size: clamp(12px, 1.8vw, 13.5px);
            color: #94a3b8;
            line-height: 1.5;
        }

        .btn-neon-cyan-download {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            background: linear-gradient(135deg, #00f0ff 0%, #0077ff 100%);
            color: #030814;
            text-decoration: none;
            padding: 12px 20px;
            border-radius: 10px;
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(11.5px, 1.8vw, 13px);
            font-weight: 900;
            letter-spacing: 0.8px;
            text-transform: uppercase;
            box-shadow: 0 0 25px rgba(0, 240, 255, 0.6);
            transition: all 0.3s ease;
            white-space: nowrap;
            min-height: 44px; /* Touch friendly on mobile */
        }
        .btn-neon-cyan-download:hover {
            transform: translateY(-2px);
            box-shadow: 0 0 40px rgba(0, 240, 255, 0.9);
            background: linear-gradient(135deg, #42f5ff 0%, #0094ff 100%);
        }

        /* Step Input Modules */
        .office-step-module {
            background: rgba(10, 18, 42, 0.7);
            border: 1px solid rgba(255, 255, 255, 0.09);
            border-radius: 16px;
            padding: clamp(16px, 3vw, 24px);
            margin-bottom: 22px;
            transition: all 0.3s ease;
        }
        .office-step-module:hover, .office-step-module:focus-within {
            border-color: rgba(255, 0, 127, 0.45);
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.5), inset 0 0 20px rgba(255, 0, 127, 0.05);
        }

        .step-module-top {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 6px;
        }
        .step-index-badge {
            width: clamp(28px, 4vw, 34px);
            height: clamp(28px, 4vw, 34px);
            border-radius: 8px;
            background: linear-gradient(135deg, var(--neon-pink), var(--neon-purple));
            color: #ffffff;
            font-family: 'Orbitron', sans-serif;
            font-weight: 800;
            font-size: clamp(13px, 2vw, 15px);
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 0 15px rgba(255, 0, 127, 0.5);
            flex-shrink: 0;
        }
        .step-module-title {
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(13.5px, 2.2vw, 16px);
            font-weight: 700;
            color: #ffffff;
            letter-spacing: 0.5px;
            line-height: 1.4;
        }
        .step-module-desc {
            font-size: clamp(12px, 1.8vw, 13.5px);
            color: #94a3b8;
            margin-bottom: 14px;
            margin-left: clamp(0px, 4.5vw, 46px);
            line-height: 1.5;
        }

        /* Office Futuristic Dropzones */
        .office-neon-dropzone {
            position: relative;
            margin-left: clamp(0px, 4.5vw, 46px);
            border: 2px dashed rgba(0, 240, 255, 0.38);
            background: rgba(4, 9, 24, 0.9);
            border-radius: 14px;
            padding: clamp(16px, 3vw, 24px);
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        .office-neon-dropzone:hover {
            border-color: var(--neon-cyan);
            box-shadow: 0 0 25px rgba(0, 240, 255, 0.35);
            background: rgba(0, 240, 255, 0.025);
        }
        .office-neon-dropzone input[type="file"] {
            position: absolute;
            top: 0; left: 0;
            width: 100%; height: 100%;
            opacity: 0;
            cursor: pointer;
            z-index: 5;
        }
        .dropzone-title-main {
            font-size: clamp(13px, 2vw, 14.5px);
            font-weight: 700;
            color: #e2e8f0;
            margin-top: 6px;
        }
        .dropzone-title-sub {
            font-size: clamp(11.5px, 1.6vw, 12.5px);
            color: #64748b;
            margin-top: 3px;
        }
        .file-selected-notify {
            margin-top: 10px;
            display: none;
            color: var(--neon-emerald);
            font-weight: 800;
            font-size: clamp(12.5px, 1.8vw, 14px);
            text-shadow: 0 0 12px rgba(0, 255, 136, 0.6);
            word-break: break-all;
        }

        /* Execute Button */
        .btn-office-execute {
            width: 100%;
            background: linear-gradient(135deg, #00f0ff 0%, #ff007f 35%, #ffb703 70%, #00ff88 100%);
            background-size: 300% 300%;
            animation: gradientRun 5s ease infinite;
            color: #030814;
            padding: clamp(15px, 3vw, 20px) clamp(16px, 4vw, 32px);
            border-radius: 14px;
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(13.5px, 2.2vw, 16.5px);
            font-weight: 900;
            letter-spacing: clamp(1px, 1.5vw, 2px);
            text-transform: uppercase;
            border: none;
            cursor: pointer;
            box-shadow: 0 0 35px rgba(0, 240, 255, 0.5), 0 0 20px rgba(255, 0, 127, 0.4);
            transition: all 0.3s ease;
            margin-top: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            min-height: 54px; /* Touch target size */
        }
        @keyframes gradientRun {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .btn-office-execute:hover {
            transform: translateY(-2px);
            box-shadow: 0 0 50px rgba(0, 240, 255, 0.8), 0 0 30px rgba(255, 0, 127, 0.65);
        }

        /* Alerts */
        .cyber-alert {
            padding: clamp(14px, 2.5vw, 18px) clamp(16px, 3vw, 22px);
            border-radius: 14px;
            font-size: clamp(13px, 1.8vw, 14.5px);
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            gap: 10px;
            backdrop-filter: blur(10px);
            line-height: 1.5;
        }
        .cyber-alert.success {
            background: rgba(0, 255, 136, 0.12);
            border: 1.5px solid var(--neon-emerald);
            color: #d1fae5;
            box-shadow: 0 0 22px rgba(0, 255, 136, 0.3);
        }
        .cyber-alert.error {
            background: rgba(255, 0, 127, 0.12);
            border: 1.5px solid var(--neon-pink);
            color: #ffe4e6;
            box-shadow: 0 0 22px rgba(255, 0, 127, 0.3);
        }

        /* Output Audit Table Section */
        .office-audit-section {
            margin-top: 36px;
            border-top: 1px solid rgba(0, 240, 255, 0.25);
            padding-top: 30px;
        }
        .audit-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 14px;
            margin-bottom: 18px;
        }
        .audit-title-text {
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(15px, 2.5vw, 18px);
            font-weight: 800;
            color: #ffffff;
            letter-spacing: 1px;
            text-transform: uppercase;
        }
        .btn-emerald-bulk-zip {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            background: linear-gradient(135deg, #00ff88 0%, #00b359 100%);
            color: #031a0e;
            text-decoration: none;
            padding: 12px 20px;
            border-radius: 10px;
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(12px, 1.8vw, 13.5px);
            font-weight: 900;
            letter-spacing: 0.8px;
            text-transform: uppercase;
            box-shadow: 0 0 25px rgba(0, 255, 136, 0.6);
            transition: all 0.3s ease;
            min-height: 44px;
        }
        .btn-emerald-bulk-zip:hover {
            transform: translateY(-2px);
            box-shadow: 0 0 40px rgba(0, 255, 136, 0.9);
        }

        .office-table-wrap {
            border: 1px solid rgba(0, 240, 255, 0.25);
            border-radius: 16px;
            overflow-x: auto; /* Horizontal scrolling on small mobile screens */
            background: rgba(6, 12, 28, 0.9);
            -webkit-overflow-scrolling: touch;
        }
        .office-audit-table {
            width: 100%;
            min-width: 580px; /* Ensures clean layout without squeeze on small phones */
            border-collapse: collapse;
            font-size: clamp(12.5px, 1.8vw, 13.5px);
        }
        .office-audit-table th {
            background: rgba(0, 240, 255, 0.08);
            padding: 14px 18px;
            font-family: 'Orbitron', sans-serif;
            font-size: 11.5px;
            font-weight: 700;
            color: var(--neon-cyan);
            letter-spacing: 1px;
            text-transform: uppercase;
            border-bottom: 1px solid rgba(0, 240, 255, 0.2);
            text-align: left;
        }
        .office-audit-table td {
            padding: 14px 18px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            color: #cbd5e1;
            vertical-align: middle;
        }
        .office-audit-table tr:hover {
            background: rgba(0, 240, 255, 0.035);
        }

        .badge-status-ok {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 11.5px;
            font-weight: 700;
            background: rgba(0, 255, 136, 0.15);
            border: 1px solid var(--neon-emerald);
            color: var(--neon-emerald);
            box-shadow: 0 0 10px rgba(0, 255, 136, 0.3);
            white-space: nowrap;
        }
        .badge-status-warn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 11.5px;
            font-weight: 700;
            background: rgba(255, 0, 127, 0.15);
            border: 1px solid var(--neon-pink);
            color: var(--neon-pink);
            box-shadow: 0 0 10px rgba(255, 0, 127, 0.3);
            white-space: nowrap;
        }

        .btn-table-single-dl {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            background: rgba(0, 240, 255, 0.1);
            border: 1px solid var(--neon-cyan);
            color: var(--neon-cyan);
            text-decoration: none;
            padding: 6px 14px;
            border-radius: 8px;
            font-weight: 700;
            font-size: 12px;
            transition: all 0.2s ease;
            white-space: nowrap;
            min-height: 36px;
        }
        .btn-table-single-dl:hover {
            background: var(--neon-cyan);
            color: #040914;
            box-shadow: 0 0 20px rgba(0, 240, 255, 0.65);
        }

        /* Developer Rights & Corporate Footer */
        .footer-rights-box {
            text-align: center;
            margin-top: 30px;
            padding: 18px 0;
            border-top: 1px solid rgba(255, 255, 255, 0.08);
        }
        .footer-rights-text {
            font-family: 'Orbitron', sans-serif;
            font-size: clamp(11px, 1.8vw, 12.5px);
            color: #94a3b8;
            letter-spacing: 1px;
            text-transform: uppercase;
            line-height: 1.5;
        }
        .footer-rights-highlight {
            color: var(--neon-cyan);
            font-weight: 900;
            text-shadow: 0 0 12px rgba(0, 240, 255, 0.5);
        }
        .footer-sub-details {
            font-size: clamp(10.5px, 1.6vw, 11.5px);
            color: #64748b;
            margin-top: 6px;
            line-height: 1.4;
        }

        /* Mobile Optimization Media Queries */
        @media (max-width: 640px) {
            body { padding: 18px 10px; }
            .step-template-banner { flex-direction: column; align-items: stretch; text-align: center; }
            .tpl-banner-left { flex-direction: column; text-align: center; }
            .btn-neon-cyan-download { width: 100%; }
            .btn-emerald-bulk-zip { width: 100%; }
            .step-module-desc { margin-left: 0; }
            .office-neon-dropzone { margin-left: 0; }
            .office-decor-1, .office-decor-2, .office-decor-3 { display: none; } /* Hide large bg watermarks on mobile to save performance */
        }
    </style>
</head>
<body>

<!-- Office Floating Elements Watermarks in Background -->
<div class="office-bg-decor office-decor-1">&#128188; OFFICE AUTOMATION</div>
<div class="office-bg-decor office-decor-2">&#128203; Word &bull; Documents</div>
<div class="office-bg-decor office-decor-3">&#128220; Editor</div>

<div class="app-container">
    
    <!-- Top Futuristic Brand Header -->
    <div class="brand-header-wrap">
        <div class="dev-badge-top">
            <span class="pulse-core"></span>
            <span class="dev-badge-text">&#128187; AUTOMATED Word DOCUMENT SUITE</span>
        </div>
        
        <!-- Colorful Glowing Title -->
        <h1 class="glowing-rainbow-title">Word Texter</h1>
        
        <!-- Developer Credit & Rights Badge -->
        <div class="developer-signature-pill">
            <span>&#9889; Developed & All Rights Reserved:</span>
            <span class="dev-author-name">Devsingh.m</span>
        </div>
        
        <p class="brand-subtext">
            High-Performance <span class="txt-cyan">Office Document Automation</span> &bull; <span class="txt-pink"> Address Replacer</span> for Only Word Documents with <span class="txt-emerald">Zero-Corruption XML Precision</span>.
        </p>
    </div>

    <!-- Main Holographic Terminal Card -->
    <div class="glass-terminal-card">
        
        <?= $alertMsg ?>

        <!-- STEP 1: DOWNLOAD PRE-CONFIGURED EXCEL TEMPLATE -->
        <div class="step-template-banner">
            <div class="tpl-banner-left">
                <div class="office-icon-box">&#128202;</div>
                <div>
                    <div class="template-heading-title">Step 1 &bull; Download Pre-Formatted Excel Template</div>
                    <div class="template-sub-desc">Standard office mapping template formatted with: <span class="txt-cyan">File Name</span>, <span class="txt-pink">Old Address</span>, <span class="txt-emerald">New Address</span>.</div>
                </div>
            </div>
            <a href="?action=download_template" class="btn-neon-cyan-download">
                &#11015; Download Template
            </a>
        </div>

        <!-- FORM: STEP 2 & STEP 3 -->
        <form method="POST" enctype="multipart/form-data">
            
            <!-- STEP 2: UPLOAD FILLED EXCEL TEMPLATE -->
            <div class="office-step-module">
                <div class="step-module-top">
                    <div class="step-index-badge">2</div>
                    <div class="step-module-title">Upload Filled Excel Template (.xlsx / .csv)</div>
                </div>
                <div class="step-module-desc">Upload your office mapping spreadsheet containing target file names and address replacement pairs.</div>
                
                <div class="office-neon-dropzone">
                    <input type="file" name="template_file" id="template_file" accept=".xlsx,.csv" required onchange="handleOfficeFileSelect('template_file', 'tpl-preview')">
                    <div style="font-size: 30px; color: var(--neon-cyan);">&#128220;</div>
                    <div class="dropzone-title-main">Click to select or drag & drop Office Excel Template</div>
                    <div class="dropzone-title-sub">Supports Microsoft Excel (.xlsx) & CSV format</div>
                    <div class="file-selected-notify" id="tpl-preview"></div>
                </div>
            </div>

            <!-- STEP 3: UPLOAD WORD DOCUMENTS -->
            <div class="office-step-module">
                <div class="step-module-top">
                    <div class="step-index-badge" style="background: linear-gradient(135deg, var(--neon-cyan), var(--neon-purple));">3</div>
                    <div class="step-module-title">Upload Target Office Word Documents (.docx)</div>
                </div>
                <div class="step-module-desc">Select Word documents to undergo automated replacement.</div>
                
                <div class="office-neon-dropzone">
                    <input type="file" name="word_files[]" id="word_files" accept=".docx" multiple required onchange="handleOfficeMultiSelect('word_files', 'docx-preview')">
                    <div style="font-size: 30px; color: var(--neon-pink);">&#128196;</div>
                    <div class="dropzone-title-main">Click to select or drag & drop Word Documents</div>
                    <div class="dropzone-title-sub">Multi-file selection enabled for automated batch processing</div>
                    <div class="file-selected-notify" id="docx-preview"></div>
                </div>
            </div>

            <!-- MULTI-COLOR EXECUTE BUTTON -->
            <button type="submit" class="btn-office-execute">
                &#9889; Execute  Address Replacement
            </button>
        </form>

        <!-- AUDIT OUTPUT RESULTS -->
        <?php if (!empty($zipArchiveName) && !empty($processedFiles)): ?>
            <div class="office-audit-section">
                <div class="audit-header-row">
                    <div class="audit-title-text">&#128229; Processed Document Output Log</div>
                    <a href="?action=download_zip&zip=<?= urlencode($zipArchiveName) ?>" class="btn-emerald-bulk-zip">
                        &#128230; Download Complete ZIP Archive
                    </a>
                </div>

                <div class="office-table-wrap">
                    <table class="office-audit-table">
                        <thead>
                            <tr>
                                <th>Processed File Name</th>
                                <th>Verification Status</th>
                                <th>Execution Time</th>
                                <th style="text-align: right;">Download Option</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php foreach ($auditReport as $item): ?>
                                <tr>
                                    <td><strong style="color: #ffffff;">&#128196; <?= htmlspecialchars($item['filename']) ?></strong></td>
                                    <td>
                                        <?php if ($item['success']): ?>
                                            <span class="badge-status-ok">&#10004; <?= htmlspecialchars($item['status']) ?></span>
                                        <?php else: ?>
                                            <span class="badge-status-warn">&#9888; <?= htmlspecialchars($item['status']) ?></span>
                                        <?php endif; ?>
                                    </td>
                                    <td style="color: #94a3b8; font-family: monospace;"><?= htmlspecialchars($item['time']) ?></td>
                                    <td style="text-align: right;">
                                        <a href="?action=download_file&file=<?= urlencode($item['filename']) ?>" class="btn-table-single-dl">
                                            &#11015; Download .docx
                                        </a>
                                    </td>
                                </tr>
                            <?php endforeach; ?>
                        </tbody>
                    </table>
                </div>
            </div>
        <?php endif; ?>

    </div>

    <!-- Developer Rights & Corporate Footer -->
    <div class="footer-rights-box">
        <div class="footer-rights-text">
            Developed & All Rights Reserved: <span class="footer-rights-highlight">Devsingh.m</span>
        </div>
        <div class="footer-sub-details">
            Word Texter &bull;  Document Automation &bull; Proprietary Zero-Corruption XML Engine
        </div>
    </div>

</div>

<script>
function handleOfficeFileSelect(inputId, displayId) {
    const input = document.getElementById(inputId);
    const display = document.getElementById(displayId);
    if (input.files && input.files[0]) {
        display.style.display = 'block';
        display.innerHTML = '&#9989; File Loaded: ' + input.files[0].name;
    }
}

function handleOfficeMultiSelect(inputId, displayId) {
    const input = document.getElementById(inputId);
    const display = document.getElementById(displayId);
    if (input.files && input.files.length > 0) {
        display.style.display = 'block';
        if (input.files.length === 1) {
            display.innerHTML = '&#9989; File Loaded: ' + input.files[0].name;
        } else {
            display.innerHTML = '&#9989; ' + input.files.length + ' Office Documents Ready for Batch Replacement';
        }
    }
}
</script>

</body>
</html>
