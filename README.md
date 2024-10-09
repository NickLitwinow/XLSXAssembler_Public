<h2 align="center">
  XLSX Assembler – ETL Tool for Merging Excel Data
</h2>
<div align="center">
  <img alt="Demo" src="./assets/Demo.png" />
</div>

<br/>

<center>

[![forthebadge](https://forthebadge.com/images/badges/built-with-love.svg)](https://forthebadge.com) &nbsp;
[![forthebadge](https://forthebadge.com/images/badges/made-with-python.svg)](https://forthebadge.com) &nbsp;
[![forthebadge](https://forthebadge.com/images/badges/open-source.svg)](https://forthebadge.com) &nbsp;
![GitHub Repo stars](https://img.shields.io/github/stars/NickLitwinow/XLSXAssembler_Public?color=blue&logo=github&style=for-the-badge) &nbsp;
![GitHub forks](https://img.shields.io/github/forks/NickLitwinow/XLSXAssembler_Public?color=blue&logo=github&style=for-the-badge)

</center>

## Built With

This project was built using these technologies.

- Python
- Airflow
- Cron
- Redis
- Pandas
- Openpyxl
- PyQT5
- Docker

## Features

**🚀 Efficient ETL Process**

Automates the extraction, transformation, and loading (ETL) of data from multiple Excel files using Airflow.\
(Only specific excel structure)

**📊 Advanced Data Processing**

Leverages the power of Pandas and Openpyxl for fast and accurate data reading, processing, and styling.

**💻 Intuitive GUI with PyQt5**

Includes a user-friendly graphical interface for selecting files and tracking real-time progress.

**⚡ Performance Optimization**

Optimized for reduced system load and faster data processing using Redis, ensuring efficient handling of large datasets.

## Getting Started

Prerequisites:
- `Python` and `Docker` installed on your machine

## 🛠 Installation and Setup Instructions

1. Clone the repository:
`git clone https://github.com/NickLitwinow/XLSXAssembler_Public.git`

2. Navigate into the `src` directory `cd src/`

4. Run the ETL client:
`python app.py`

5. Build the Docker image:
`docker build . --tag extending_airflow:latest`

6. Run `docker-compose up -d` command to start the services.

The PyQt5 GUI will launch, where you can select multiple Excel files and begin the ETL process.
*Runs the app in the development mode.*

## Usage Instructions Example

1. In the ETL client click `Add File` button and select files from the `example files` (You can add them again later if you want so)
2. (Optional) To remove a file from selected, click on it's path (element) in the black selection window. Click `Remove File` to remove the file.
3. Click `Merge Files` to name the output file and choose it's destination. The ETL process will start afterwards.
4. To view the Airflow Dag process, open `http://localhost:8080/home` with login: `airflow` and password: `airflow`.
5. To view the Radis database:
- Open `http://localhost:8001/`
- Accept "EULA and Privacy Settings"
- Click `I already have a database`
- Click `Connect to a Radis Database` with Host: `redis`, Port: `6379`, Name: `redis-local`
- Click `ADD REDIS DATABASE`
- Select the `redis-local` database
 
### Show your support

Give a ⭐ if you like this project!
