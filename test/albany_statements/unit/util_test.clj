(ns albany-statements.unit.util-test
  (:require [clojure.test :refer [deftest testing is]]
            [albany-statements.util :refer :all]))

(deftest to-col-test
  (testing "to-col behaves itself."
    (is (= (to-col  0)  :A))
    (is (= (to-col 25)  :Z))
    (is (= (to-col 26) :AA))))

(deftest tocurrency-test
  (testing "to-currency produces correct output."
    (is (= (tocurrency 4) "4.00"))
    (is (= (tocurrency 4.3) "4.30"))
    (is (= (tocurrency 4.124) "4.12"))
    (is (= (tocurrency 4.126) "4.13"))
    (is (= (tocurrency -4.124) "-4.12"))))

(deftest member-cols-test
  (testing "member-cols produces correct map."
    (is (= (member-cols 4 [:des :cost]) {:E :des :F :cost}))
    (is (= (member-cols 6 [:a :b :c]) {:G :a :H :b :I :c}))
    (is (= (member-cols 0 [:map-starts-with-A]) {:A :map-starts-with-A}))))

(deftest try-get-integer-test
  (testing "try-get-integer recognises integers."
    (is (= (try-get-integer 7) 7))
    (is (= (try-get-integer 3.14) 3.14))
    (is (= (try-get-integer 987.0) 987))
    (is (= (try-get-integer 12.0) 12))
    (is (= (try-get-integer 9876.54) 9876.54))))

(deftest essential-cases-test
  (testing "Essential cases arithmetic works."
    (is (= (essential-cases  6  6) 1))
    (is (= (essential-cases 10  5) 2))
    (is (= (essential-cases  3  6) 1/2))
    (is (= (essential-cases  4  6) 2/3))))

(deftest format-description-test
  (testing "Formatting of the description."
    (is (= (format-description "[Brand] Product (size) ORG")
           [:span [:em "Brand"] [:b " Product (size) "] "ORG"]))))

(deftest roundup-test
  (testing "roundup behaves correctly."
    (is (= (roundup 1) 1))
    (is (= (roundup 12.0) 12))
    (is (= (roundup 15.7) 16))))
