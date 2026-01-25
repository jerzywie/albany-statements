(ns albany-statements.unit.balance-test
  (:require [dk.ative.docjure.spreadsheet :refer [select-columns load-workbook select-sheet]]
            [clojure.test :refer [deftest testing is]]
            [albany-statements.balance :refer :all]
            [mockery.core :refer [with-mock with-mocks]]))

(deftest member-balance-data-tests
  (testing "keyworize-bal"
    (is (= (keywordize-bal {:name "Amount" :subname nil :val 97.05 :key "owed"})
           {:owed {:name "Amount" :subname nil :val 97.05}})))

  (let [expected-member-balance-maps
        {:owed {:name "Amount" :subname nil :val 97.05}
         :bf {:name "Balance" :subname "sub-bal" :val 14.58}}]
    (testing "get-member-balance"
      (let [sel-cols-result
            [{:name "Amount" :subname nil :val 97.05 :key "owed"}
             {:name "Balance" :subname "sub-bal" :val 14.58 :key "bf"}
             {:name "ANOther" :subname "ANOther" :val nil :key nil}]]
        (with-mock _
          {:target :dk.ative.docjure.spreadsheet/select-columns
           :return sel-cols-result}
          (let [g-m-b-result (get-member-balance :PersonA _)]
            (is (= (count sel-cols-result) 3))
            (is (= (count g-m-b-result) 2))
            (is (= g-m-b-result expected-member-balance-maps))))))
    (testing "bal-item"
      (is (= (bal-item :owed expected-member-balance-maps) ["Amount" 97.05])))))

(deftest member-data-test
  (testing "member-data"
    (is (= (take 2 member-data) '([:PersonD 10] [:PersonI 20])))))



