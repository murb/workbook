# Changelog

### 0.9.0

* Dropped inheritance from Array and borrowed from the Enumerable module instead (this might break some existing implementations; please create pull request if you miss Array like functionality)
* Fix: ODS: Currency formatted cell now no longer returns String

### 0.8.1

* Adopted [Standard](https://github.com/testdouble/standard) for code formatting

### 0.8.0

* Dropped support for ruby < 2.3.0
* Added support for 2.6.0
* Stripping strings before calling to_sym (possibly breaking, as symbols can be used to address columns)
* Started maintaining a changelog
