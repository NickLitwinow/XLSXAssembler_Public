from datetime import datetime
from airflow.decorators import dag, task
from excel_pipe.read_data import read_data
from excel_pipe.assemble_data import assemble_data
from excel_pipe.insert_excel_data import insert_excel_data
import pickle
from confluent_kafka import Producer, Consumer, KafkaException

# Kafka producer setup
producer_conf = {
    'bootstrap.servers': 'kafka:9092',
    'message.max.bytes': 1000000000}
producer = Producer(producer_conf)

# Kafka consumer setup
consumer_conf = {
    'bootstrap.servers': 'kafka:9092',
    'group.id': 'airflow_group',
    'auto.offset.reset': 'latest'
}
consumer = Consumer(consumer_conf)

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

        # Data serialization for storing
        serialized_data = pickle.dumps(all_sheets_data)

        # Sending data to Kafka
        print('Sending data to Kafka')
        producer.produce('read_data_topic', value=serialized_data)
        producer.flush()
        print('Sent data to Kafka')

        msg = 'r_data done' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        return msg

    @task()
    def a_data(msg):
        print(msg)

        # Consuming data from Kafka
        print('Consuming data from Kafka')
        consumer.subscribe(['read_data_topic'])
        msg = consumer.poll(1.0)
        if msg is None:
            raise KafkaException("No message received from Kafka")
        else:
            print('Consumed data from Kafka')

        all_sheets_data = pickle.loads(msg.value())

        dfs = assemble_data(all_sheets_data)

        # Data serialization for storing
        serialized_dfs = pickle.dumps(dfs)

        # Sending data to Kafka
        print('Sending data to Kafka')
        producer.produce('assemble_data_topic', value=serialized_dfs)
        producer.flush()
        print('Sent data to Kafka')

        msg = 'a_data done' + datetime.strftime(datetime.now(), '%m/%d/%Y %H:%M:%S')
        return msg

    @task()
    def i_data(msg):
        print(msg)

        # Consuming data from Kafka
        print('Consuming data from Kafka')
        consumer.subscribe(['assemble_data_topic'])
        msg = consumer.poll(1.0)
        if msg is None:
            raise KafkaException("No message received from Kafka")
        else:
            print('Consumed data from Kafka')

        dfs = pickle.loads(msg.value())

        insert_excel_data(dfs, output_file)

    # Последовательность выполнения тасков
    r = r_data(file_paths=file_paths, output_file=output_file)
    a = a_data(r)
    i_data(a)

# Запуск DAG
excel_dag = combine_excel_sheets(file_paths=[], output_file='')
