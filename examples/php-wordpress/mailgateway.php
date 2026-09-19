<?php
/**
 * MailGateway client for PHP / WordPress.
 *
 * Drop into a mu-plugin or theme functions.php. Define the two constants in wp-config.php:
 *   define('MG_URL',     'https://apps.central-office.info/mailgateway/api/mail/send');
 *   define('MG_API_KEY', '.....');
 *
 * Example:
 *   mg_send([
 *     'to'       => 'member@example.org',
 *     'from'     => 'noreply@pasmeeting.org',
 *     'fromName' => 'PAS Meeting',
 *     'subject'  => 'Your receipt',
 *     'htmlBody' => '<p>Thanks! Receipt attached.</p><img src="cid:logo">',
 *     'attachments' => [
 *        mg_attach_file('/var/www/receipts/12345.pdf', 'Receipt-12345.pdf'),
 *        mg_attach_file(__DIR__ . '/logo.png', 'logo.png', 'logo'),   // inline image
 *     ],
 *   ]);
 */

function mg_attach_file(string $path, ?string $name = null, ?string $contentId = null): array
{
    $item = [
        'name'          => $name ?: basename($path),
        'contentBase64' => base64_encode(file_get_contents($path)),
    ];
    if ($contentId) {
        $item['contentId'] = $contentId;
        $item['isInline']  = true;
    }
    return $item;
}

function mg_attach_bytes(string $bytes, string $name, ?string $contentId = null): array
{
    $item = ['name' => $name, 'contentBase64' => base64_encode($bytes)];
    if ($contentId) {
        $item['contentId'] = $contentId;
        $item['isInline']  = true;
    }
    return $item;
}

/**
 * @param array $message  keys: to, cc, bcc, replyTo, from, fromName, subject, htmlBody, textBody, attachments, appName
 * @return array ['ok' => bool, 'status' => int, 'requestId' => string, 'error' => string|null, 'raw' => string]
 */
function mg_send(array $message): array
{
    $message['appName'] = $message['appName'] ?? (defined('WP_HOME') ? WP_HOME : 'php');

    $response = wp_remote_post(MG_URL, [
        'timeout' => 300,
        'headers' => [
            'Content-Type' => 'application/json; charset=utf-8',
            'X-API-Key'    => MG_API_KEY,
        ],
        'body' => wp_json_encode($message),
    ]);

    if (is_wp_error($response)) {
        return ['ok' => false, 'status' => 0, 'requestId' => '', 'error' => $response->get_error_message(), 'raw' => ''];
    }

    $status = (int) wp_remote_retrieve_response_code($response);
    $raw    = (string) wp_remote_retrieve_body($response);
    $json   = json_decode($raw, true) ?: [];

    return [
        'ok'        => $status >= 200 && $status < 300 && !empty($json['success']),
        'status'    => $status,
        'requestId' => $json['requestId'] ?? '',
        'error'     => $json['error'] ?? null,
        'raw'       => $raw,
    ];
}

/**
 * Optional: route ALL WordPress mail (wp_mail) through the gateway.
 * Uncomment to enable. Attachments passed to wp_mail() as file paths are picked up automatically.
 */
// add_filter('pre_wp_mail', function ($null, array $atts) {
//     $headers = is_array($atts['headers'] ?? null) ? $atts['headers'] : explode("\n", (string) ($atts['headers'] ?? ''));
//     $isHtml = false;
//     foreach ($headers as $h) if (stripos($h, 'content-type:') === 0 && stripos($h, 'text/html') !== false) $isHtml = true;
//
//     $attachments = [];
//     foreach ((array) ($atts['attachments'] ?? []) as $path) if (is_readable($path)) $attachments[] = mg_attach_file($path);
//
//     $result = mg_send([
//         'to'          => is_array($atts['to']) ? implode(';', $atts['to']) : $atts['to'],
//         'subject'     => $atts['subject'],
//         'htmlBody'    => $isHtml ? $atts['message'] : null,
//         'textBody'    => $isHtml ? null : $atts['message'],
//         'attachments' => $attachments,
//     ]);
//     if (!$result['ok']) error_log('MailGateway: ' . $result['error'] . ' (request ' . $result['requestId'] . ')');
//     return $result['ok'];
// }, 10, 2);
