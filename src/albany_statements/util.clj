(ns albany-statements.util
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [garden.stylesheet]
            [clojure.string :as s]))

(def brand-delim-start "[")
(def brand-delim-end "]")
(def org-marker "ORG")

;;This was obtained from Stackoverflow
(defn to-col [num]
  (loop [n num s ()]
    (if (> n 25)
      (let [r (mod n 26)]
        (recur (dec (/ (- n r) 26)) (cons (char (+ 65 r)) s)))
      (keyword (apply str (cons (char (+ 65 n)) s))))))

(defn tocurrency
  "Format a number for currency display (2 digits. Parentheses when negative)"
  [v]
  (format "%.2f" (double v)))

(defn member-cols
  "Produce the map of columns for a given member"
  [index keys]
  (zipmap  (vec  (map to-col (range index (+ (count keys) index)))) keys ))

(defn try-get-integer
  "Try to return an integer. Return passed value otherwise."
  [v]

  (let [bigint-v (/ (int (* v 100)) 100)
        int-v (int bigint-v)]
    (if (= bigint-v int-v) int-v v)))

(defn essential-cases
  "Calculate the quantity of an order item as a number of Essential cases."
  [des-albany-qty albany-units-per-case]
  (let [int?-des-albany-qty (try-get-integer des-albany-qty)
        int?-albany-units-per-case (try-get-integer albany-units-per-case)]
    (try-get-integer (/ int?-des-albany-qty int?-albany-units-per-case))))

(defn format-description
  "Try to extract product brand and apply emboldening to product name"
  [rawdesc]
  (if (s/starts-with? rawdesc brand-delim-start)
    (let [end-brand (s/index-of rawdesc brand-delim-end)
          org-pos (s/index-of rawdesc org-marker)
          prod-name (if org-pos
                      (subs rawdesc (inc end-brand) org-pos)
                      (subs rawdesc (inc end-brand)))
          brand-name (subs rawdesc 1 end-brand)]
      [:span
       [:em brand-name]
       [:b prod-name]
       (if org-pos org-marker)])
    [:b rawdesc]))
