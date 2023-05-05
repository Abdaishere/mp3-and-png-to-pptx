# .mp3 and .png into .pptx
convert .mp3 and .png into pptx slide

## Installation

install [python-pptx](https://pypi.org/project/python-pptx/) to install pptx.

```bash
pip install python-pptx
```
python 3.8.8

## Usage

```python
from pptx import Presentation
from pptx.util import Inches
from os import listdir
from os.path import isfile, join

prs = Presentation()

# insert .mp3 and .png files in src folder 
# the .mp3 file must be the same order name as the .png file
src = '.\\src\\'
fileNames = [f for f in listdir(src) if isfile(join(src, f))]
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.
