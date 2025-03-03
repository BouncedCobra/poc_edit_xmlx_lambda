import json
import boto3
import io
from openpyxl import load_workbook

s3 = boto3.client('s3')

def lambda_handler(event, context):
    try:
        # Mostrar lo que tiene el bucket
        bucket = 'poc-edit-excel-s'
        response = s3.list_objects_v2(Bucket=bucket)
        for obj in response.get('Contents', []):
            print(obj['Key'])

        # Obtener el archivo Excel de S3
        key = 'plantilla.xlsx'
        response = s3.get_object(Bucket=bucket, Key=key)
        data = response['Body'].read()

        # Leer la hoja 'TITULAR' del archivo Excel
        with io.BytesIO(data) as input_stream:
            workbook = load_workbook(input_stream)
            titular_sheet = workbook['TITULAR']
            operatividad_sheet = workbook['OPERATIVIDAD']

            # Verificar si la celda B18 es parte de un rango combinado
            for merged_cell in titular_sheet.merged_cells.ranges:
                if 'B18' in merged_cell:
                    # Modificar la celda superior izquierda del rango combinado
                    titular_sheet[merged_cell.coord.split(':')[0]].value = 'Sergio Trujillo'
                    break
                else:
                    # Si no es parte de un rango combinado, modificar la celda directamente
                    titular_sheet['B18'] = 'Sergio Trujillo'

            # Modificar la celda B18 de la hoja 'OPERATIVIDAD'
            for merged_cell in operatividad_sheet.merged_cells.ranges:
                if 'B7' in merged_cell:
                    operatividad_sheet[merged_cell.coord.split(':')[0]].value = 'Sergio Garcia'
                    break
                else:
                    operatividad_sheet['B7'] = 'Sergio Garcia'

        # Guardar el archivo modificado en S3
        modified_key = 'modified_plantilla.xlsx'
        with io.BytesIO() as output_stream:
            workbook.save(output_stream)
            output_stream.seek(0)
            s3.put_object(Bucket=bucket, Key=modified_key, Body=output_stream.read())

        return {
            'statusCode': 200,
            'body': json.dumps('Hello from Lambda!')
        }

    except Exception as e:
        print(e)
        return {
            'statusCode': 500,
            'body': json.dumps('Error')
        }