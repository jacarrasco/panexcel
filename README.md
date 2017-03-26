PanExcel
=======================

Tool to Export rules and objects from Panorama Palo Alto Networks into an excel file

## Options

| Argument  |  Description |
|---|---|
|  -f   | Firewall name. If none is selected it will produce an excel of all firewalls  |
|  -v | Enable if xml file has Virtual Systems instead of Device Groups |
|  -c | Introduce a concrete config file. By default config.xml is read  |
| -r  |  Introduce a concrete rulename to obtain a report for just one rule |
|  -o |  Introduce a concrete object to obtain a report of all rules associated to that object |
|  -e |  Select a concrete output filename |


## Usage examples



## Dependencies
Python with the following modules:

* lxml
* xlsxwriter
* argparse
