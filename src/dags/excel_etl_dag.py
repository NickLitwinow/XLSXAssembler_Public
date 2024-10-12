# Imports the datetime class for timestamp generation.
from datetime import datetime
# Imports Airflow decorators to define DAGs and tasks.
from airflow.decorators import dag, task
# Import custom ETL functions
from excel_pipe.extract_data import read_data
from excel_pipe.transform_data import assemble_data
from excel_pipe.load_data import insert_excel_data
# Imports the pickle module for data serialization (converting data structures to byte streams).
import pickle
# Imports the redis module for interacting with the Redis in-memory data store.
import redis

# Establishes a connection to a Redis server running on the default host (redis) and port (6379).
r = redis.Redis(host='redis', port=6379)

def redis_set(msg, data):
    """
    Stores data in Redis under a key prefixed with 'X_data' and the current timestamp.
    r.set(msg, data): Sets the key-value pair using the Redis SET command.
    """

    r.set(msg, data)

def redis_get(msg):
    """
    Retrieves data from Redis using the provided key. Deletes the key after retrieval.
    res = r.get(msg): Fetches the data associated with the key using the Redis GET command.
    r.delete(msg): # Removes the key from the Redis store after retrieval.
    return res: Returns the retrieved data.
    """

    res = r.get(msg)
    r.delete(msg)
    return res


'''
Defines an Airflow DAG named combine_excel_sheets with the following properties:
default_args: A dictionary containing default arguments for tasks within the DAG.
'owner': Assigns ownership of the DAG to "Nikita Litvinov".
'retries': Disables retries for tasks within this DAG (set to 0).
start_date: Sets the DAG start date to the current time.
catchup=False: Prevents the DAG from running for past dates on manual execution.
schedule_interval=None: Specifies that the DAG is triggered manually.
'''

default_args = {
    'owner': 'Nikita Litvinov',
    'retries': 0,
}

@dag(dag_id='combine_excel_sheets',
     default_args=default_args,
     start_date=datetime.now(),
     catchup=False,
     schedule_interval=None)
def combine_excel_sheets(file_paths, output_file):
    """
    Defines a function representing the DAG's execution flow.
    - file_paths: A list of paths to Excel files to be processed.
    - output_file: The path to the final output Excel file.
    """

    @task()
    def extract_data(file_paths, output_file):
        """
        def extract_data(file_paths, output_file): Reads data from the provided Excel files using the read_data function.
        - all_sheets_data = read_data(file_paths, output_file): Calls the custom read_data function to retrieve data from the files.
        - serialized_data = pickle.dumps(all_sheets_data): Serializes the retrieved data using pickle.dumps.
        - redis_set(msg, serialized_data): Stores the serialized data in Redis using the redis_set function.
        - return msg: Returns the message used for Redis key generation.
        """

        all_sheets_data = read_data(file_paths, output_file)
        serialized_data = pickle.dumps(all_sheets_data)

        print('Set data to Redis')
        msg = 'extract_data ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        redis_set(msg, serialized_data)
        print('Set completed')

        return msg

    @task()
    def transform_data(msg):
        """
        def transform_data(msg): Assembles data from the combined source.
        all_sheets_data = pickle.loads(redis_get(msg)): Deserializes the data retrieved from Redis using pickle.loads.
        serialized_data = pickle.dumps(assemble_data(all_sheets_data)): Serializes the assembled data.
        redis_set(msg, serialized_data): Stores the serialized assembled data in Redis.
        return msg: Returns the message used for Redis key generation.
        """

        print('Get data from Redis')
        all_sheets_data = pickle.loads(redis_get(msg))
        print('Get completed')

        serialized_data = pickle.dumps(assemble_data(all_sheets_data))

        print('Sending data to Redis')
        msg = 'transform_data ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        redis_set(msg, serialized_data)
        print('Sent data to Redis')

        return msg

    @task()
    def load_data(msg, output_file):
        """
        def load_data(msg, output_file): Inserts the processed data into the output Excel file.
        dfs = pickle.loads(redis_get(msg)): Deserializes the data retrieved from Redis using pickle.loads.
        insert_excel_data(dfs, output_file): Runs the assembled data processing function insert_excel_data function.
        """

        print('Consuming data from Redis')
        dfs = pickle.loads(redis_get(msg))
        print('Consumed data from Redis')

        insert_excel_data(dfs, output_file)

    # Sequence of task execution
    msg = extract_data(file_paths=file_paths, output_file=output_file)
    msg = transform_data(msg)
    load_data(msg, output_file=output_file)


# Launches DAG with parameters of selected XLSX files and output file path
excel_dag = combine_excel_sheets(file_paths=[], output_file='')
