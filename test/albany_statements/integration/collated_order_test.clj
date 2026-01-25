(ns albany-statements.integration.collated-order-test
  (:require [dk.ative.docjure.spreadsheet :refer [select-sheet]]
            [albany-statements.core :refer :all]
            [clojure.test :refer [use-fixtures deftest testing is]]))

(def re-bound-out "test/outputs/test.txt")
(def test-spreadsheet "dev-resources/test-order.xlsm")

(defn clean-up-files [f]
  (clojure.java.io/delete-file re-bound-out true)
  (f)
  (clojure.java.io/delete-file re-bound-out true))

(use-fixtures :each clean-up-files)

(deftest closed-tests
  (testing "sheet-status"
    (binding [*out* (new java.io.FileWriter re-bound-out)]
      (is (thrown? IllegalStateException (check-sheet-status false true)))
      (check-sheet-status false false)
      (is (clojure.string/includes? (slurp  re-bound-out) "'Open'" )))))

(deftest spreadsheet-tests
  (testing "spreadsheet read gets expected data"
    (let [sheet-data (get-spreadsheet-data  test-spreadsheet true)]
      (is (clojure.set/subset? #{:wb :revision :order-sheet} (set (keys sheet-data))))
      (is (= (select-sheet "Collated Order" (:wb sheet-data)) (:order-sheet sheet-data)))
      (let [persona-order (get-member-order [:PersonA 15] (:order-sheet sheet-data))]
        (is (= (count (:PersonA persona-order)) 4))
        (is (= (-> persona-order :PersonA first keys set) (set '(:description :code :case-size :singles-per-case :unit-cost :vat-amount :memdes :memcost))))))))
