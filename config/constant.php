<?php
/**
 * Config : constant variables
 * This file is used to handle custom contant values.
 *
 * Do not use helper function in config file. eg. url(), asset()
 *
 * @author JQ Gan <gan.jinqu@i-serve.com.my>
 */

define('HEADER_USERTYPE', 'User-Type');
define('HEADER_USERAGENT', 'User-Agent');
define('HEADER_USERID', 'User-Id');
define('HEADER_SOURCE', 'Source');
define('DATABASE_PATH', base_path('database'));
define('DATE_TIME', 'Y-m-d H:i:s');
define('DATE', 'Y-m-d');

define('SOURCE_WEB', 'bmZ0Y2hhcl9haXJkcm9wX3dlYjozMjMzZTQ1MWVmZjBiOWVmMTZhNTljNzI1YmFkZTlmNg==');
define('SOURCE_ANDROID', 'bmZ0Y2hhcl9haXJkcm9wX2FuZHJvaWQ6NjI1OGFjNzliYWE0MTdlODZlZDE0OWQxYWIxN2JiMzQ==');
define('SOURCE_IOS', 'bmZ0Y2hhcl9haXJkcm9wX2lvczoxZWQ1MGU0Nzk5MDZlYmU1NjY5MjdjZjRjOTViYzllMQ==');

return [
    // 'source_authentication' => [SOURCE_WEB, SOURCE_ANDROID, SOURCE_IOS],
    'source_authentication' => [SOURCE_WEB => 'web', SOURCE_ANDROID => 'android', SOURCE_IOS => 'ios'],
    'special_character' => '"%*;<>?^`{|}~\\\'#=&',
    'day_in_seconds' => 86400,
];
