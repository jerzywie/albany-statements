(ns albany-statements.config
  (:require [clojure.edn :as edn]))

(def config-data-string (slurp "resources/config.data"))

(defn config-data-edn [edn-string] (edn/read-string edn-string))

(def config-data (config-data-edn config-data-string))
