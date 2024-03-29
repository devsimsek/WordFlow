# WordFlow - Another Awesome Static Site Generator

## About This Project

WordFlow is a static site generator built by devsimsek using Python.

Its main goal is to be sourced using DOCX and Markdown files.

Be aware that I am still developing this project, and currently it's far from a stable release, so do expect bugs and
errors.

## Installation

Clone this repo and run main.py

```sh
git clone https://github.com/devsimsek/WordFlow . && python main.py
```

WordFlow will automaticly call initializer for you to configure.

## Roadmap

* [x] Add category integration
* [ ] Work on json api generation. Order content by their type and all.
* [ ] Create more friendly cli
* [ ] Migrate to the markdown to html generator library (undependent)
* [ ] Theme partialization
* [ ] Automatic version checker & updater
* [ ] Integrate plugin development
* [ ] Optimize code
* [ ] Create markdown editor which has the ability to directly generate site. (As a notepad)

## Changelog

* Sept 17, 2023 - Removed category integration. Optimized code. Tests will begin soon.
* Jan 31, 2023 - Added category integration. Now you can just create subdirectories as a category.
* Jan 30, 2023 - Added Proper readme & Started integration of the markdown language & Made public json api
* Jan 9, 2023 - Optimization around the docx text renderer
* Jan 8, 2023 - Optimized homepage generation
* Jan 6, 2023 - Optimizations around docx generator (mostly images)
* Jan 4, 2023 - Optimizations around the docx generator & Added homepage generation & Added table to html generation
* Jan 2, 2023 - Initializer Optimization
* Jan 1, 2023 - Initial Release

## Possible Bugs

- [ ] (Investigating) Initializer generated error when using utf-8 characters
	- Details: UnicodeDecodeError: 'utf-8' codec can't decode byte 0xc4 in position 16: invalid continuation byte

## Active Users

Thanks to all users of this project.
| Name | Site Address | Category | Source |
| ---- | ------------ | -------- | ---- |
| Simsek's Notes | [simseks.github.io](https://simseks.github.io) | Blog | docx (Microsoft Word) |
| Lessons by devsimsek | [lessons.smsk.me](http://lessons.smsk.me) | Blog | md (Markdown)|

If you are using WordFlow and want to share your site here please open a pull request.

## Contributing

Contributions are welcome. Just fork this repository, make your changes and create pull request!

## Authors

- [devsimsek](https://beta.smsk.me)

## License

[MIT License](https://devsimsek.mit-license.org)