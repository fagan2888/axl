# axl

Python excel add-ins, made easy!

# axl requirements

- pywin32
- pandas

# System Requirements

- Microsoft Excel
- Microsoft Windows
- Anaconda or Miniconda

# Installing

In your conda environment, run `conda install -c mcg axl` to install the
package.

# Adding the Excel Plugin

In the environment root folder, find the axl excel plugin at `<anaconda path>\envs\<axl env path>\axl.xlam`. 
Add this file to your Excel addins under File -> Options -> Add-ins. The basic 
axl functionality is now available through the Excel-Python API.

# Adding Libraries

To add more functions to the Excel, first ensure that the libraries you are
going to use are installed into the conda environment are installed through pip
or conda.

axl will add a folder in your environment's Tools directory where you will list
which functions you want included. It is located at `<anaconda path>\envs\<axl env path>\Tools\axl`.

In this folder, for each library you want to use, add a file named
`imports.<library name>`. In this file, add all of your import statements.
axl will parse this file and add all of those imports to the Excel function
namespace.

You may need to close and open Excel for the functions to populate.

# Using axl

All the libraries you imported will be avaiable through the Excel formulas bar
with special syntax for python.

For example, if you have added `from math import factorial` in an imports file, you can run the
factorial function on a cell like so: `=P("factorial",B1)`.
