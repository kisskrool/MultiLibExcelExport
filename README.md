# MultiLibExcelExport

This library is intended to offer compatibility from the PHP language to multiple
exporting Excel libraries:
- PHPExcel 1.7.9
- LibXL (C++) using the php_excel C extension
- Spread::WriteExcel 2.40 (Perl) using the php_perl C extension
- php_writeexcel 0.3.0

The interest is to be able to switch between libraries for example if they are 
not maintained anymore or for performance reasons. Optimization has been kept
in mind while developing as it was required to export big data volumetries.

All was tested on PHP 5.3, 5.4, 5.5 and should be working on most recent versions.
It has also been tested in multithreading environment using the php_pthreads C
extension.

It's composed of a common object CellMatrix which is used to represent an Excel
worksheet containing Cell objects.

Facades are used to abstract workbooks, worksheets and styles. And adapters are 
used for each library which follow a common pattern defined by an interface.

Each library usually contains a common set of functionnalities which are in the
interface, but they can also present particularities that are not common and can
define their own functions in adapters.

The project has begun to be developed following the functionnalities of php_writeexcel
which is not maintainted anymore. Some code has been modeled or taken on
the PHPExcel project which (the author will pardon me) presented too much overhead
during execution but had also lot of really useful functionnalities.