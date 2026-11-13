<?php
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    exit(0);
}

$configFile = __DIR__ . '/config.json';

// Function to get local LAN IP
function getLocalIp() {
    $ip = 'localhost';
    if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
        $output = shell_exec('ipconfig');
        if ($output && preg_match('/IPv4 Address[ .]*:[ ]*([0-9.]+)/i', $output, $matches)) {
            $ip = $matches[1];
        }
    } else {
        $output = shell_exec("hostname -I | awk '{print $1}'");
        if ($output) {
            $ip = trim($output);
        }
    }
    return $ip;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $raw = file_get_contents('php://input');
    $data = json_decode($raw, true);

    if (isset($data['gasUrl'])) {
        $cleanUrl = trim($data['gasUrl']);
        $config = [
            'gasUrl' => $cleanUrl,
            'updated_at' => date('Y-m-d H:i:s')
        ];
        file_put_contents($configFile, json_encode($config, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));

        echo json_encode([
            'success' => true,
            'gasUrl' => $cleanUrl,
            'message' => 'บันทึกการเชื่อมต่อไปยังเซิร์ฟเวอร์เรียบร้อยแล้ว ทุกอุปกรณ์ (มือถือ, iPad, PC) จะเชื่อมต่อชีตนี้โดยอัตโนมัติ'
        ]);
        exit;
    } else {
        echo json_encode(['success' => false, 'error' => 'Missing gasUrl parameter']);
        exit;
    }
}

// GET request: Load current configuration & server info
$gasUrl = '';
if (file_exists($configFile)) {
    $config = json_decode(file_get_contents($configFile), true);
    if (isset($config['gasUrl'])) {
        $gasUrl = trim($config['gasUrl']);
    }
}

$serverIp = getLocalIp();
$port = isset($_SERVER['SERVER_PORT']) && $_SERVER['SERVER_PORT'] != '80' ? ':' . $_SERVER['SERVER_PORT'] : '';
$mobileUrl = "http://{$serverIp}{$port}/0002/index.html";

echo json_encode([
    'success' => true,
    'gasUrl' => $gasUrl,
    'server_ip' => $serverIp,
    'mobile_url' => $mobileUrl
]);
