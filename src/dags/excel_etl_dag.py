from datetime import datetime
from airflow.decorators import dag, task
from excel_pipe.read_data import read_data
from excel_pipe.assemble_data import assemble_data
from excel_pipe.insert_excel_data import insert_excel_data
import pickle
import redis

r = redis.Redis(host='redis', port=6379)
def redis_set(msg, data):
    r.set(msg, data)

def redis_get(msg):
    res = r.get(msg)
    r.delete(msg)
    return res


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

    @task()
    def r_data(file_paths, output_file):
        all_sheets_data = read_data(file_paths, output_file)

        serialized_data = pickle.dumps(all_sheets_data)

        # Sending data to Redis
        print('Set data to Redis')
        msg = 'r_data ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        redis_set(msg, serialized_data)
        print('Set completed')

        return msg

    @task()
    def a_data(msg):

        # Receive data from Redis
        print('Get data from Redis')
        all_sheets_data = pickle.loads(redis_get(msg))
        print('Get completed')

        serialized_data = pickle.dumps(assemble_data(all_sheets_data))

        # Sending data to Redis
        print('Sending data to Redis')
        msg = 'a_data ' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        redis_set(msg, serialized_data)
        print('Sent data to Redis')

        return msg

    @task()
    def i_data(msg, output_file):

        # Consuming data from Redis
        print('Consuming data from Redis')
        dfs = pickle.loads(redis_get(msg))
        print('Consumed data from Redis')

        insert_excel_data(dfs, output_file)

    # Последовательность выполнения тасков
    r_msg = r_data(file_paths=file_paths, output_file=output_file)
    a_msg = a_data(r_msg)
    i_data(a_msg, output_file=output_file)


# Запуск DAG
excel_dag = combine_excel_sheets(file_paths=[], output_file='')
