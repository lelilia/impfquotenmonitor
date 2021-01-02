# Impfquotenmonitor

This will generate a small website on the overall national progress of vaccinations against Covid-19 in Germany and will compare the total number to a German city of a similar size.

You can see this at https://www.johl.io/impfdaten/

The data is pulled from the Robert Koch-Institut that publishes data on vaccination daily except on weekends. City population numbers are queried from Wikidata.

## Getting Started

Install the project from GitHub. You need Python3.

### Prerequisites

Make sure you have the necessary libraries installed. Enter the `impfquotenmonitor` directory.

Run this:

```
pip3 install -r requirements.txt
```

### Running the program

Run this:

```
python3 impfquotenmonitor.py
```

This will generate an `index.html` file. Feel free to copy it on a webserver to make it available to the world.
