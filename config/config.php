<?php if (!defined('TL_ROOT')) die('You can not access this file directly!');

/**
 * Contao Open Source CMS
 *
 * Copyright (c) 2005-2015 Leo Feyer
 *
 * @license LGPL-3.0+
 */

if (extension_loaded("xml") && extension_loaded("zip"))
{
	// Add Hook for EFG Excel export
	$GLOBALS['TL_HOOKS']['efgExportXls'][] = array('EfgExcelExport', 'export');
}

?>