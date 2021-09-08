# albany-statements

Albany order forms and statements derived from xls

## Usage

Run the project directly, via `:main-opts` (`-m albany-statements.core`):

    $ clojure -M:run-m
    
Build an uberjar:

    $ clojure -X:uberjar

Run the uberjar:

    $ java -jar albany-statements.jar [args]

## Options

    Run the project without arguments to see this usage message:
    

    Usage: program-name [options] spreadsheet-name order-date-id

    Options:
      -o, --output-type OUTPUT-TYPE      Type of output: order-forms | statements
      -c, --coordinator COORDINATOR      Coordinator name (required for output-type=order-forms)
      -v, --version VERSION          :d  Version of order-form: draft  | final
      -h, --help

    Both spreadsheet-name and order-date-id must be supplied

## License

Copyright © 2013 - 2021 Jerzy

Distributed under the Eclipse Public License, the same as Clojure.
