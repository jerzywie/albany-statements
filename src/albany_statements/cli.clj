(ns albany-statements.cli
  (:require [dk.ative.docjure.spreadsheet :refer [select-columns load-workbook select-sheet]]
            [clojure.string :as string]
            [clojure.tools.cli :refer [parse-opts]]))

(def opt-sep " | ")

(def output-types {:o "order-forms" :s "statements"})
(def versions {:d "draft" :f "final"})

(defn version-tostring [version-key] (version-key versions))

(defn list-options [valid-map] (string/join opt-sep (vals valid-map)))

(defn keywordise-option [opt] (keyword (string/lower-case (subs opt 0 1))))

(defn keyword-validator [valid-map]
  [(fn [v] (contains? (set (keys valid-map)) v))
   (str "must be one of: " (list-options valid-map))])

(def cli-options
  [;; First three strings describe a short-option, long-option with optional
   ;; example argument description, and a description. All three are optional
   ;; and positional.
   ["-o" "--output-type OUTPUT-TYPE" (str "Type of output: " (list-options output-types))
    :parse-fn keywordise-option
    :validate (keyword-validator output-types)]
   ["-c" "--coordinator COORDINATOR" "Coordinator name (required for output-type=order-forms)"]
   ["-v" "--version VERSION" "Version of order-form: draft  | final"
    :parse-fn keywordise-option
    :validate (keyword-validator versions)
    :default :d]
   ["-n" "--no-closed-check" "Do not check the 'Closed' status before processing."
    :default false]
   ["-h" "--help"]])


(defn usage
  ([options-summary message]
   (->> ["Albany Order-Form/Statement generator."
         "Usage: program-name [options] spreadsheet-name order-date-id"
         ""
         "Options:"
         options-summary
         ""
         message]
        (string/join \newline)
        (println)))

  ([output-type]
   (usage output-type nil)))

(defn process-args [args]
  (let [{:keys [options arguments errors summary]} (parse-opts args cli-options)]
    (cond
      (:help options) (usage summary)
      (not= (count arguments) 2) (usage summary "Both spreadsheet-name and order-date-id must be supplied")
      errors (usage errors)
      :else
      (let [[spreadsheet-name order-date]  arguments]
        (assoc options
               :spreadsheet-name spreadsheet-name
               :order-date order-date
               :summary summary)))))
