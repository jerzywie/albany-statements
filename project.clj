(defproject albany-statements "0.3.2"
  :description "Albany order forms and statements based on order xls"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [dk.ative/docjure "1.6.0"]
                 [hiccup "1.0.4"]
                 [garden "0.1.0-beta6"]]
  :profiles {:uberjar {:aot :all}}
  :uberjar-name "albany-statements.jar"
  :main albany-statements.core
  :aot [albany-statements.config])
