
# {a11ytables2}

<!-- badges: start -->
[![Project Status: WIP – Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)
<!-- badges: end -->

This package is an experimental work-in-progress successor to [{a11ytables}](https://co-analysis.github.io/a11ytables/).

The plan is for {a11ytables2} to:

* be built on [{openxlsx2}](https://github.com/JanMarvin/openxlsx2/) instead of {openxlsx}
* take advantage of a 'YAML blueprint' system for specifying a11ytables-class objects (allowing for arbitrary pre-table metadata, easier addition of multiple tables per sheet, etc)
* take advantage of tools like {cli} and {fs} for better messaging and file handling

You should use {a11ytables} until {a11ytables2} is stable (which may never happen).
