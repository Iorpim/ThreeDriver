## Threedriver
----------------
### Features
Threedriver is a script for mapping and matching OneDrive/SharePoint directories. It is Selenium based and will first crawl, and then, if matchers are specified, download the files looking for the specified regular expressions.

It can currently search in `.docx`, `.pptx`, `.xlsx`,  `.msg`, `.pdf`, and any other pure text files as `.txt`, `.csv`, and so on.
File format detection is not extension, but header reliant, avoiding most false positives/false negatives.


### Setup
As any Selenium project it requires a working browser and it's webdriver. For Chromium it can be found here: https://chromedriver.chromium.org/downloads
For some distros it can also be found in the package manager, please refer to your prefered search engine on the best source for your system.

More details on Firefox setup can be found further into this page.


##### Using enviroment variables:

Linux 
```
pip install -r requirements.txt
export WEBDRIVER=$WEBDRIVER_PATH
./threedriver.py -h
```

Windows 
```
pip install -r requirements.txt
set WEBDRIVER=$WEBDRIVER_PATH
./threedriver.py -h
```

##### Using local executable file
If the `WEBDRIVER` enviroment variable is not set by default the script looks for `.\chromedriver.exe` on Windows and `./chromedriver` on any other system.

```
pip install -r requirements.txt
./threedriver.py -h
```


### Usage
```
python threedriver.py -h
usage: threedriver.py [-h] [--verbose] [--max-depth N] [--blacklist FILE_PATH] [--delay T] [--json OUTPUT_FILE] [--matcher REGEX_STRING] [--quiet]
                      [--max-size SIZE_LIMIT] [--keep-files PERMANENT_DOWNLOAD_FOLDER] [--match-binary-files] [--disable-msg-recursion]
                      [--extension-blacklist .EXTENSION] [--load-fs FS_JSON] [--office-365 LOGIN_URL] [--user USER] [--password PASSWORD] [--proxy PROXY_URI]
                      [--firefox]

optional arguments:
  -h, --help            show this help message and exit
  --verbose, -v
  --max-depth N, -d N
  --blacklist FILE_PATH, -e FILE_PATH
  --delay T, -t T
  --json OUTPUT_FILE, -o OUTPUT_FILE
  --matcher REGEX_STRING, -m REGEX_STRING
  --quiet, -q
  --max-size SIZE_LIMIT, -L SIZE_LIMIT
  --keep-files PERMANENT_DOWNLOAD_FOLDER, -k PERMANENT_DOWNLOAD_FOLDER
  --match-binary-files, -A
  --disable-msg-recursion, -f
  --extension-blacklist .EXTENSION, -x .EXTENSION
  --load-fs FS_JSON, -l FS_JSON
  --office-365 LOGIN_URL, -b LOGIN_URL
  --user USER, -u USER
  --password PASSWORD, -p PASSWORD
  --proxy PROXY_URI, -P PROXY_URI
  --firefox, -F
```

The verbose flag can be repeated for increased output.
`--max-depth` represents how many layers into the file system the script will crawl.

`--blacklist` referers to files/folders to be ignored from crawling/matching. As `--matchers` this is an incremental flag.

`--delay` is the base delay unit in seconds, without reliable ways to define if the page has fully loaded, or complete an action, most operations rely on a multiple of the base delay unit. 
The faster the internet connection and it's reliability, the smaller this value can be, and vice versa. 
The default value is 1 second, increase in case of frequent errors during processing.

`--json` allows the output to be exported into a json file, the results will be split in two, `JSON_fs.json` and `JSON_matches.json`.

`--matcher` is a required flag for matching, the application lacks any default matchers. This flag is incremental, meaning that for multiple matchers the following syntax should be used: `./threedriver.py … -m "example" -m "second example"`.
The matchers are passed to Python's `re.compile` function, and should support any standard Python regular expression, although without proper support for capture groups.

`--quiet` disables the printing of the file system an matches to console.
Usage of this flag requires the setting of the `--output` flag.

`--max-size` limits the size of files to be matched, as these are downloaded to disk first, it could be a problem if large files are encountered during execution.
This flag can use any human readable format. Ex. `20KB`, `20000`, `50MB`
(Note that it uses decimal, not binary conversion.)
Default: unlimited

`--keep-files` keeps the files after matching, allowing files to reviewed manually later. Takes a folder as argument, to which files will be moved after download.

`--match-binary-files`, by default unknown binary files are skipped, this flag forces the matching on these too.

`--disable-msg-recursion` disables matching in .msg attachment files. By default the script will recurse into the the attachments, proceeding further if other .msg files or documents are found.

`--extension-blacklist` defines which extensions to ignore during matching. If the extension is matched the file won't even be downloaded. 
Note this is a simple matching of the end of the file name, a too open filter could have unwanted results, for example `txt` would also match `exampletxt` and not only `example.txt`.
As with `--blacklist` and `--matchers` this is also an incremental flag, so for multiple extensions the flag should be passed multiple times.

`--load-fs` points to a `JSON_fs.json` file which can be used to skip the crawling step in case it has alreaby been completed before. If `--json` is set this file is exported before matching, meaning this will still be created even in case of any errors during later execution stages.

`office-365`  takes the SharePoint URL to be used, must be used for proper execution if not using simple OneDrive.

`--user`, username to be used for login. If blank will prompt during execution.

`--password`, password to be used for login. Will also be prompted during execution if blank.

`--proxy` complete proxy URL as to be passed to webdriver. Ex: `socks5://127.0.0.1:8080`

`--firefox` untested option allowing usage with Firefox instead of Chrome/Chromium. More details on the current limitations for this option can be found below.


### Firefox support

Firefox is not currently supported for matching operations, or usage with a proxy, for that some adaptations need to be made in the `__init__` function. Namely the proxy support, the way it finds the driver executable, and to set the default download folder.


### Current development state
Although fully functional for OneDrive, SharePoint uses significantly different URLs and HTML code, even while looking very similar, and thus the script is currently unable to extract and process SharePoint data.

##### TODO
Lines with pending features can be located by the tag `TODO`, while lines with broken features can be found looking for the tag `BROKEN`, these lines also include a brief description of the problem, and the requirements for a fix.