#from pandas.core.indexing import _LocationIndexer
import logging
import os

import psycopg2 as ps
from psycopg2.extensions import register_adapter, AsIs
from sqlalchemy import create_engine
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import datetime as dt
import pyodbc
import sqlite3
import urllib
import warnings
from dotenv import load_dotenv

warnings.simplefilter(action='ignore', category=FutureWarning)

load_dotenv()

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format=os.getenv("LOG_FORMAT", "%(asctime)s | %(levelname)s | %(message)s"),
    datefmt=os.getenv("LOG_DATEFMT", "%Y-%m-%d %H:%M:%S")
)

logger = logging.getLogger(__name__)

#import hanconfig

register_adapter(np.int64, AsIs)
register_adapter(np.int32, AsIs)
register_adapter(np.float64, lambda x: AsIs(float(x)))

def castAsString(data):

    """Casts the specified data piece as a string with correct formatting for dates."""

    if isinstance(data, int) or isinstance(data, np.int64) or isinstance(data, float):

        return str(data)

    elif isinstance(data, str):

        return data

    elif isinstance(data, pd._libs.tslibs.timestamps.Timestamp):

        return datetime.strftime(data, '%d/%m/%Y')

    else:

        return str(data)

def castAsBoolean(data):

    """Turns binary data into Boolean."""

    if data == 1:

        return True

    elif data == 0:

        return False

    # Use pd.isna() for NaN detection - np.nan != np.nan in NumPy 2.x
    elif data == '' or data is None or pd.isna(data):

        return ''

    else:

        raise Exception("Tried to cast a value that was not 1, 0 or null as boolean. Please check your parameters.")

def readData(filepath, skiprows=0, sheet_name=None):

    """Catch-all to turn most filetypes into dataframes."""

    if filepath.find('.xls') >= 0:

        # If the excel has more than 1 sheet, please specify the sheet name

        if sheet_name:

            return pd.read_excel(filepath, skiprows=skiprows, sheet_name=sheet_name)

        else:

            return pd.read_excel(filepath, skiprows=skiprows)

    elif filepath.find('.csv') >= 0:

        return pd.read_csv(filepath)

    elif filepath.find('.txt') >= 0:

        return pd.read_csv(filepath)

    else:

        raise Exception(filepath + " error: File cannot be read by pandas.")

def dataType(data):

    """Assesses the inherent type of one piece of data and returns the data type name it corresponds to, which are utilised in singleType function."""

    # Check for pandas NA type first
    if isinstance(data, pd._libs.missing.NAType):

        return 'Null'

    # Check for NaT (Not a Time) - must be before date check
    elif isinstance(data, pd._libs.tslibs.nattype.NaTType):

        return 'Date'

    # Use pd.isna() for NaN detection - handles np.nan, None, pd.NA, pd.NaT
    # Must check this BEFORE isinstance(data, float) since np.nan is a float
    elif pd.isna(data):

        return 'Null'

    # Check for explicit null-like string values
    elif data in ['#N/A', '']:

        return 'Null'

    # Check bool BEFORE int because bool is a subclass of int in Python
    elif isinstance(data, (bool, np.bool_)):

        return 'Boolean'

    # Check for integers (including numpy integer types)
    elif isinstance(data, (int, np.integer)):

        return 'Integer'

    # Check for floats (including numpy float types)
    elif isinstance(data, (float, np.floating)):

        return 'Numeric'

    elif isinstance(data, str):

        if data.lower() == 'true' or data.lower() == 'false':

            return 'Boolean'

        else:

            return 'Varchar'

    elif isinstance(data, (pd._libs.tslibs.timestamps.Timestamp, dt.date)):

        return 'Date'

    else:

        return 'Varchar'

def singleType(arr, flavour):

    """A column may have more than one data type - this function aims to minimise data type conflict by setting an entire column to specific data types in hierarchal fashion."""

    # If you need to add more SQL flavours, add them to the following object. If you need to add more data types, e.g. BLOB, bigint, JSON, add them to the object.

    # The general format is: 'generic_data_type':'specific_type_identifier' for each SQL flavour, e.g. Varchar in Postgres is called text in Access. I have defaulted to Postgres as its syntax is more common

    # Note for SQLite date doesn't exist, needs to be text however the date field must be a string

    flavour_types = {'postgres':
                        {'Varchar':'Varchar',
                        'Boolean':'Boolean',
                        'Integer':'Integer',
                        'Numeric':'Numeric',
                        'Date':'Date'},
                    'access':
                        {'Varchar':'text',
                        'Boolean':'yesno',
                        'Integer':'integer',
                        'Numeric':'numeric',
                        'Date':'date'},
                    'sqlite':
                        {'Varchar':'text',
                        'Boolean':'integer',
                        'Integer':'integer',
                        'Numeric':'real',
                        'Date':'date'}
    }

    if 'Varchar' in arr:

        data_type = 'Varchar'

    elif 'Boolean' in arr:

        data_type = 'Boolean'

    elif 'Numeric' in arr:

        data_type = 'Numeric'

    elif 'Date' in arr:

        data_type = 'Date'

    elif 'Integer' in arr:

        data_type = 'Integer'

    else:

        data_type = 'Varchar'

    return flavour_types[flavour][data_type]

def postgresDate(date):

    pg_date_string = datetime.strftime(date, '%Y-%m-%d')

    return pg_date_string

def booleanToString(data, flavour):

    flavour_bools = {'postgres':
                        {'true':'true',
                        'false':'false'
                        },
                    'access':
                        {'true':'-1',
                        'false':'0'},
                    'sqlite':
                        {'true':1,
                        'false':0}
    }

    if isinstance(data, str):

        if data.lower() == 'true':

            return flavour_bools[flavour]['true']
            
        elif data.lower() == 'false':

            return flavour_bools[flavour]['false']

        else:

            return flavour_bools[flavour]['false']

    elif data:

        return flavour_bools[flavour]['true']

    else:

        return flavour_bools[flavour]['false']

def cleanHeaders(dataframe, table_name, flavour):

    """Removes/modifies conflicting headers from a dataframe as certain phrases such as date, text, value will confuse some SQL flavours. This function is enabled by default meaning often no changes are necessary to the underlying data."""

    # SQLite/Access don't rely on schemas

    if flavour in ['access', 'sqlite']:

        if table_name.find('real_') >= 0 or table_name.find('temp_') >= 0:

            table_prefix = table_name[5:]

        elif table_name.find('raw') >= 0:

            table_prefix = table_name[5:]

        else:

            table_prefix = table_name

    elif flavour == 'postgres':

        if table_name.find('real.') >= 0 or table_name.find('temp.') >= 0 or table_name.find('data.') >= 0:

            table_prefix = table_name[5:]

        elif table_name.find('raw.') >= 0:

            table_prefix = table_name[4:]

        else:

            table_prefix = table_name

    else:

        table_prefix = table_name.replace('.', '_') if table_name.find('.') >= 0 else table_name

    # Just add restricted names to the list if you find an exception.

    restricted_column_names = ['date', 'currency', 'text', 'primary', 'value']

    for column in dataframe.columns:

        clean_column = column.lower()

        if clean_column in restricted_column_names:

            dataframe.columns = dataframe.columns.str.replace(column, table_prefix + '_' + clean_column)

        clean_column = clean_column.replace(' ', '_')
        clean_column = clean_column.replace('-', '_')

        dataframe.columns = dataframe.columns.str.replace(column, clean_column)

    return dataframe

class sqliteDatabase(object):

    """Creates and connects to a SQLite database, which is useful for storing in Dropbox and can be queried using a database explorer.
    
    To set a custom db_path, call sqliteDatabase(db_path='your_path_here') - this is not necessary if using the default directory as it is built in."""

    def __init__(self, db_path="default"):

        self.user = os.getlogin()

        self.flavour = 'sqlite'

        if db_path == "default":

            self.db_path = "C:\\Users\\{username}\\HAN ETF Dropbox\\Shared\\Database\\Alternative Databases\\HANetf Database.db".format(username=self.user)

        else:

            self.db_path = db_path

        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()

    def disconnect(self):

        """Disconnects from the database - only necessary when running multiple scripts in a row to clear memory. Not explicitly needed in most cases."""

        self.conn.close()

    def query(self, query, write=False, commit=False, parse_dates=[]):

        if write:

            self.cursor.execute(query)
            logger.debug("Query executed with write permission. If you wanted to request data instead, please use write=False.")

        else:

            sql_data = pd.read_sql(query, self.conn, parse_dates=parse_dates)

            return sql_data

        if commit:

            self.conn.commit()
            logger.debug("Committed.")

    def createTable(self, dataframe, table_name, append=False, clean_dataframe=True, log=False, add_id=False):

        """Creates a table from a dataframe. Schemas do not exist."""

        if clean_dataframe:

            dataframe = dataframe.replace("'","''",regex=True)
            dataframe = dataframe.fillna('')

            cleanHeaders(dataframe, table_name, 'sqlite')

        column_names = dataframe.columns

        if append:

            try:

                placeholder = pd.read_sql("select * from {table} limit 1".format(table=table_name), self.conn)
                logger.info(f"Appending into {table_name}")

            except:

                logger.info(f"{table_name} did not exist, creating.")
                append = False

        if not append:

            try:

                logger.debug(f"Attempting to drop {table_name}")

                drop_query = 'drop table ' + table_name

                self.cursor.execute(drop_query)

                self.conn.commit()

                logger.debug(f"Dropped {table_name}")

            except:

                logger.debug("Tried to drop a table that doesn't exist.")

                self.conn.rollback()

            data_types = []

            for col in column_names:

                column_data_types = []

                for row in range(dataframe.shape[0]):

                    data_type = dataType(dataframe[col][row])

                    column_data_types.append(data_type)

                deduped_column_types = list(set(column_data_types))
                single_column_type = singleType(deduped_column_types, 'sqlite')

                data_types.append(single_column_type)

            # Cleanup - SQLite does not accept date fields or date as a type

            for index, (col, type) in enumerate(zip(column_names, data_types)):

                logger.debug(f"{index} {col} {type}")

                if type == 'date':

                    dataframe[col] = dataframe[col].map(castAsString)
                    data_types[index] = 'text'

            column_types = [col + ' ' + col_type for col, col_type in zip(column_names,data_types)]

            table_parameters = '(' + ','.join(str(col) for col in column_types) + ')'

            query = 'create table ' + table_name + ' ' + table_parameters

            if log:

                logger.debug(query)

            self.cursor.execute(query)

            self.conn.commit()

            logger.info(f"Created {table_name}")

        injection = 'insert into ' + table_name + ' values (' + ','.join(('?') for col in column_names) + ')'

        for row in range(dataframe.shape[0]):

            data_to_inject = [dataframe[col][row] if not (isinstance(dataframe[col][row], bool) or isinstance(dataframe[col][row], np.bool_)) else booleanToString(dataframe[col][row], 'sqlite') for col in column_names]

            logger.debug(data_to_inject)

            if log:

                logger.debug(injection)
                logger.debug(data_to_inject)

            self.cursor.execute(injection, data_to_inject)

            logger.debug(f'{table_name}: Row {row} of {dataframe.shape[0]} injected.')

        logger.info(f"{dataframe.shape[0]} rows loaded into {table_name}")
        self.conn.commit()
        logger.debug("Committed.")

class accessDatabase(object):

    """Creates an Access database connection and connects automatically. No arguments explicitly needed unless you want to set a custom db_path, e.g. accessDatabase(db_path="your_path_here")
    
    A common source of errors in Access is writing SQL using standard Postgres or Oracle syntax - pyodbc will throw an invalid syntax error if you attempt this."""

    def __init__(self, db_path="default", log=False):

        self.user = os.getlogin()

        if db_path == "default":

            self.db_path = "C:\\Users\\{username}\\HAN ETF Dropbox\\Shared\\Database\\HANetf Database.accdb".format(username=self.user)

        else:

            self.db_path = db_path

        self.flavour = 'access'

        if log:

            logger.debug(f"Username: {self.user}")
            logger.debug(f"Database path: {self.db_path}")

    def connect(self):

        self.conn_str = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            "DBQ=" + self.db_path)

        self.conn_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(self.conn_str)}"
        self.engine = create_engine(self.conn_uri)

        self.conn = pyodbc.connect(self.conn_str)
        self.cursor = self.conn.cursor()

    def disconnect(self):

        """Disconnects from the database - only necessary when running multiple scripts in a row to clear memory. Not explicitly needed in most cases."""

        self.conn.close()

    def query(self, query, write=False, commit=False, parse_dates=[]):

        """Specify a query and whether write should be True or False. If write=True, the command will be executed; if write=False, a dataframe will be returned. commit=False does not guarantee that Access won't commit due to the autocommitting nature."""

        if write:

            self.cursor.execute(query)
            logger.debug("Query executed with write permission. If you wanted to request data instead, please use write=False.")

        else:

            sql_data = pd.read_sql(query, self.conn, parse_dates=parse_dates)

            return sql_data

        if commit:

            self.conn.commit()

    def createTable(self, dataframe, table_name, add_id=False, append=False, clean_dataframe=True, log=False):

        """Creates a table called table_name from a dataframe.

        Please note that append and add_id functionality has been removed for Access."""

        if clean_dataframe:

            dataframe = dataframe.replace("'","''",regex=True)
            dataframe = dataframe.fillna('')

            cleanHeaders(dataframe, table_name, 'access')

        with self.engine.connect() as sqla_conn:

            dataframe.to_sql(table_name, sqla_conn, if_exists='replace', index=False)

        logger.info(f"{dataframe.shape[0]} rows loaded into {table_name}")
        self.conn.commit()
        logger.debug("Committed.")

class postgresDatabase(object):

    """Creates a postgres connection and automatically connects.

    Requires the following input: postgresDatabase(db_name, host, port, username, password) with every entry as a string."""

    def __init__(self, db_name, host, port, username, password):

        # Engine used for creating tables in SQLAlchemy
        # Conn used for querying and upserts

        self.engine = create_engine('postgresql+psycopg2://{user}:{password}@{host}/{db_name}'.format(user=username, password=password, host=host, db_name=db_name))

        self.conn = ps.connect(dbname=db_name, host=host, port=port, user=username, password=password)
        self.cursor = self.conn.cursor()

        self.db_name = db_name
        self.host = host
        self.port = port
        self.username = username
        self.password = password

        self.flavour = 'postgres'

    def disconnect(self):
        
        """Disconnects from the database - only necessary when running multiple scripts in a row to clear memory. Not explicitly needed in most cases."""

        self.conn.close()

    def rollback(self):

        """Rolls back previous transaction."""

        self.conn.rollback()

    def commit(self):

        """Forces a commit."""

        self.conn.commit()
        logger.debug("Committed.")

    def query(self, query=None, sql_file=None, write=False, commit=True, parse_dates=[], rollback_on_fail=True, run_safely=True):

        """Specify a query and whether write should be True or False. If write=True, the command will be executed; if write=False, a dataframe will be returned.
            If write=True, then commit=True by default, so make sure to test in a dev environment first."""

        if sql_file:

            # Gather feedback on whether encoding needs to be a parameter, or utf-8 is appropriate for all scripts

            with open(sql_file, encoding='utf-8') as f:
                query = f.read()

        if run_safely:

            try:

                if write:

                    self.cursor.execute(query)
                    logger.debug("Query executed with write permission. If you wanted to request data instead, please use write=False.")

                    if commit:

                        self.conn.commit()
                        logger.debug("Committed.")

                else:

                    sql_data = pd.read_sql(query, self.conn, parse_dates=parse_dates)

                    return sql_data

            except:

                logger.error("Query failed.")

                if rollback_on_fail:

                    self.conn.rollback()
                    logger.warning("Rollback.")

                else:

                    logger.warning("No rollback.")

        else:

                if write:

                    self.cursor.execute(query)
                    logger.debug("Query executed with write permission. If you wanted to request data instead, please use write=False.")

                    if commit:

                        self.conn.commit()
                        logger.debug("Committed.")

                else:

                    sql_data = pd.read_sql(query, self.conn, parse_dates=parse_dates)

                    return sql_data         


    def createTable(self, dataframe, table_name, schema='dev', add_id=True, append=False, clean_dataframe=True, column_data_types=[], specific_column_type=None, log=False, constraint=None):

        # does forcing the specific_column_types for a non existent column break things?

        """Creates a table called table_name in the specified schema from a dataframe. As is standard postgres syntax, you must specify a schema or the table will not be created.
        
        If append is set to False, the table will be created from scratch even if it already exists, add_id will add an autoincrementing id column to the table.

        The array column_data_types is used to overwrite the 'best guess' approach to column types - you need to specify, in order, data types for EACH column if you wish to use this.

        The parameter specific_column_type lets you pass a dictionary with column names and types to replace the best guesses, useful for large tables.

        Constraint takes a constraint name as a string (e.g. unique_id) and will upsert around this. The constraint must exist prior to creating/appending for this to work."""

        if clean_dataframe:

            dataframe = dataframe.replace(pd.NA, np.nan, regex=True)

            if constraint:

                dataframe = dataframe.fillna('')

            dataframe = dataframe.replace("'","''",regex=True)
            cleanHeaders(dataframe, table_name, 'postgres')

            if 'boolean' in column_data_types or 'Boolean' in column_data_types:

                for col, type in zip(dataframe.columns, column_data_types):

                    if type in ['boolean', 'Boolean']:

                        dataframe[col] = dataframe[col].map(castAsBoolean)

            if log:

                logger.debug(dataframe)

        column_names = dataframe.columns

        if append:

            existence_query = "select not exists (select from information_schema.tables where table_schema like '" + schema + "' and table_name = '" + table_name + "')"
            stuff = pd.read_sql(existence_query, self.conn)
            create_table = pd.read_sql(existence_query, self.conn).iloc[0,0]

        else:

            drop_query = 'drop table if exists ' + schema + '.' + table_name

            self.cursor.execute(drop_query)
            self.conn.commit()
            create_table = True

        if create_table:

            if column_data_types:

                column_types = [col + ' ' + col_type for col, col_type in zip(column_names,column_data_types)]

            else:

                data_types = []

                for col in column_names:

                    column_data_types = []

                    for row in range(dataframe.shape[0]):

                        data_type = dataType(dataframe[col][row])

                        column_data_types.append(data_type)

                    deduped_column_types = list(set(column_data_types))
                    single_column_type = singleType(deduped_column_types, 'postgres')

                    data_types.append(single_column_type)

                column_types = [col + ' ' + col_type for col, col_type in zip(column_names,data_types)]
                
            if log:
                for column in column_types:
                    logger.debug(column)

            if specific_column_type:
                for column in specific_column_type:
                    try:
                        col_index = list(column_names).index(column)
                        column_types[col_index] = column + ' ' +  specific_column_type[column]
                    except:
                        pass

            if add_id:

                table_parameters = '(id serial primary key,' + ','.join(str(col) for col in column_types) + ')'

            else:

                table_parameters = '(' + ','.join(str(col) for col in column_types) + ')'

            query = 'create table ' + schema + '.' + table_name + ' ' + table_parameters

            if log:

                logger.debug(query)

            self.cursor.execute(query)

            self.conn.commit()

            logger.info(f"Created {table_name} in {schema}")

        if constraint is not None:
            logger.debug("Using psycopg2")
            table_name = schema + '.' + table_name

            column_names = dataframe.columns
            
            dataframe = dataframe.fillna('')

            for row in range(dataframe.shape[0]):

                row_of_data = [dataframe[col][row] if not (isinstance(dataframe[col][row], bool) or isinstance(dataframe[col][row], np.bool_)) else booleanToString(dataframe[col][row], 'postgres') for col in column_names]

                data_to_inject = []
                columns_to_inject = []

                for data_piece, column in zip(row_of_data, column_names):

                    if data_piece != '':
                        # Convert np.float64 to native Python float for proper SQL serialization
                        if isinstance(data_piece, np.float64):
                            data_piece = float(data_piece)
                        elif isinstance(data_piece, np.int64):
                            data_piece = int(data_piece)

                        data_to_inject.append(data_piece)
                        columns_to_inject.append(column)

                data_to_inject = data_to_inject + data_to_inject

                injection = 'insert into ' + table_name + '(' + ','.join(col for col in columns_to_inject) + ')' + ' values (' + ','.join(('%s') for col in columns_to_inject) + ') on conflict on constraint ' + constraint + ' do update set (' + ','.join(col for col in columns_to_inject) + ') = (' + ','.join(('%s') for col in columns_to_inject) + ')'
                if log:
                    logger.debug(f"{injection} {data_to_inject}")
                self.cursor.execute(injection, data_to_inject)

        else:
            logger.debug("Using SQLAlchemy")
            with self.engine.connect() as sqla_conn:

                dataframe.to_sql(table_name, sqla_conn, schema=schema, if_exists='append', index=False)

        logger.info(f"{dataframe.shape[0]} rows loaded into table {schema}.{table_name} within database: {self.db_name}")
        self.conn.commit()
        logger.debug("Committed.")


def copyDatabase(source_database, destination_database, specific_tables=[], ignore_tables=['data.holdings', 'data.sftp_data_paths', 'data.german_equity_ratio', 'data.fefundinfo', 'data.email_data_paths'], schemas_to_copy=['data', 'real'], truncate=['real.index_values'], force_schema=None, log=False):

    """Feed a source (Postgres) and destination (Postgres, Access, SQLite) connection, and mirrors all tables from schemas_to_copy over with tidied table names. Ignores tables specified in the ignore_tables array."""

    if destination_database.flavour == 'postgres':

        destination_table_query = "table_name"

    else:

        destination_table_query = "table_schema||'_'||table_name"

    schemas = "('" + "','".join(schema for schema in schemas_to_copy) + "')"

    table_list_query = '''select table_name as raw_table_name, table_schema||'.'||table_name as source_table, {destination_table_query} as destination_table, table_schema
                        from information_schema.tables
                        where table_schema in {schemas}
                        order by table_schema, table_name'''.format(schemas=schemas, destination_table_query=destination_table_query)

    tables_to_copy = source_database.query(table_list_query)

    if specific_tables:

        for raw_table_name, source_table, destination_table, schema in zip(tables_to_copy['raw_table_name'], tables_to_copy['source_table'], tables_to_copy['destination_table'], tables_to_copy['table_schema']):

            if source_table in specific_tables:
            
                source_data_type_query = """select column_name, data_type
                                        from information_schema.columns
                                        where table_schema = '{schema}' and table_name = '{raw_table_name}'""".format(schema=schema, raw_table_name=raw_table_name)

                source_data_types = source_database.query(source_data_type_query)

                parse_dates = []

                for row in range(source_data_types.shape[0]):

                    if source_data_types['data_type'][row].find('date') >= 0:

                        parse_dates.append(source_data_types['column_name'][row])

                source_dataframe = source_database.query('select * from {table}'.format(table=source_table), parse_dates=parse_dates)

                if force_schema:

                    if destination_database.flavour == 'postgres':

                        destination_database.createTable(source_dataframe, table_name=destination_table, schema=force_schema, append=False, add_id=False, log=log)

                    else:

                        destination_database.createTable(source_dataframe, destination_table, append=False, add_id=False, log=log)

                else:

                    if destination_database.flavour == 'postgres':

                        destination_database.createTable(source_dataframe, table_name=destination_table, schema=schema, append=False, add_id=False, log=log)

                    else:

                        destination_database.createTable(source_dataframe, destination_table, append=False, add_id=False, log=log)
    else:

        for raw_table_name, source_table, destination_table, schema in zip(tables_to_copy['raw_table_name'], tables_to_copy['source_table'], tables_to_copy['destination_table'], tables_to_copy['table_schema']):

            if source_table in ignore_tables:

                continue
            
            source_data_type_query = """select column_name, data_type
                                    from information_schema.columns
                                    where table_schema = '{schema}' and table_name = '{raw_table_name}'""".format(schema=schema, raw_table_name=raw_table_name)

            source_data_types = source_database.query(source_data_type_query)

            parse_dates = []

            for row in range(source_data_types.shape[0]):

                if source_data_types['data_type'][row].find('date') >= 0:

                    parse_dates.append(source_data_types['column_name'][row])

            source_dataframe = source_database.query('select * from {table}'.format(table=source_table), parse_dates=parse_dates)

            if force_schema:

                if destination_database.flavour == 'postgres':

                    destination_database.createTable(source_dataframe, table_name=destination_table, schema=force_schema, append=False, add_id=False, log=log)

                else:

                    destination_database.createTable(source_dataframe, destination_table, append=False, add_id=False, log=log)

            else:

                if destination_database.flavour == 'postgres':

                    destination_database.createTable(source_dataframe, table_name=destination_table, schema=schema, append=False, add_id=False, log=log)

                else:

                    destination_database.createTable(source_dataframe, destination_table, append=False, add_id=False, log=log)

    logger.info(f"Database successfully mirrored from {source_database.flavour} to {destination_database.flavour}.")

def getdatabaseobject(hc, database='db'):
    db = postgresDatabase(hc.cfg[database]['database'],
                          hc.cfg[database]['host'],
                          hc.cfg[database]['port'],
                          hc.cfg[database]['user'],
                          hc.cfg[database]['password'])
    return db