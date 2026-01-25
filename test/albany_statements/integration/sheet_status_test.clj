(ns albany-statements.integration.sheet-status-test
  (:require [dk.ative.docjure.spreadsheet :refer [load-workbook]]
            [albany-statements.sheet-status :refer :all]
            [clojure.test :refer [deftest testing is]]))

(deftest sheet-status-reading
  (testing "Expected status data is returned"
    (let [wb (load-workbook "dev-resources/test-order.xlsm")
          data (get-sheet-status wb)]
      (is (= (count data)) 2)
      (is (some #{:state} (keys data)))
      (is (some #{:revision} (keys data))))))

