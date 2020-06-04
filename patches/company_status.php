<?php
defined("_VALID_ACCESS") || die('Direct access forbidden');

Utils_CommonDataCommon::new_array(
    "Companies_Status",
    [
        'active_stable' => 'Aktywny stały',
        'active_unstable' => 'Aktywny - skoczek',
        'at_issue' => 'Sporny',
        'inactive_friendly' => 'Nieaktywny - mamy kontakt',
        'inactive' => 'Nieaktywny - brak kontaktu',
    ]
);

Utils_RecordBrowserCommon::new_record_field(
    'company',
    [
        'name' => _M('client status'),
        'type' => 'commondata',
        'extra' => false,
        'visible' => true,
        'required' => false,
        'position' => 28,
        'param' => "Companies_Status",
    ]
);
