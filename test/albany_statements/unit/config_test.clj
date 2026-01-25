(ns albany-statements.unit.config-test
  (:require [clojure.test :refer [deftest testing is]]
            [albany-statements.config :refer :all]))

(deftest config-data-test
  (testing "config.data is processed correctly"
    (is (= (config-data-edn "{:a [1 2 3]}") {:a [1 2 3]}))))

