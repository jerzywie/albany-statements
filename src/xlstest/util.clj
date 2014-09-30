(ns xlstest.util
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [garden.stylesheet]))

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
