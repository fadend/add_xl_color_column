# Add a column with the color per row in a spreadsheet

It's pretty easy dealing with the values from Excel files, especially since you
can convert them to CSV. However, if the color found in each row also has
meaning that you need to deal with, that can be painful.

This simple utility adds a "color" column with the first color found per row.

A better idea would probably be to simply output the color to standard out
and pair that with a utility to insert a column. That would make it more
niceable composable in the Unix way.

But this is still useful for the little project I was looking at so I'm
publishing it as is in order to help preserve the bit of knowledge I've
gained working on it.

## Usage

```
python -m venv colenv
cd colenv
bin/pip3 install git+https://github.com/fadend/add_xl_color_column
bin/python3 -m add_xl_color_column.add_xl_color_column \
  -i $HOME/Downloads/hoverflies.xlsx  -o $HOME/Downloads/hoverflies_color.xlsx 
```

## Acknowledgments

Thank you, Eric Gazoni and Charlie Clark, for the super helpful
[openpyxl](https://pypi.org/project/openpyxl/) package!

Code was formatted using [black](https://github.com/psf/black).
