(ns albany-statements.unit.core-test
  (:require [clojure.test :refer :all]
            [albany-statements.core :refer :all]))

(deftest gt-zero-test
  (testing "positive values."
    (is (true? (gt-zero 4)))
    (is (true? (gt-zero 5/3)))
    (is (true? (gt-zero 3.14159))))
  (testing "zero."
    (is (false? (gt-zero 0))))
  (testing "negative values."
    (is (false? (gt-zero -1)))
    (is (false? (gt-zero -45.67))))
  (testing "anything else."
    (is (false? (gt-zero "blah")))
    (is (false? (gt-zero {:a "A"})))))
