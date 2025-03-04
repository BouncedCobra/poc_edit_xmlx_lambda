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

        with io.BytesIO(data) as input_stream:
            workbook = load_workbook(input_stream)
            # Leer la hoja TITULAR y modificar los datos
            titular_sheet = workbook['TITULAR']

            titular_sheet = modify_excel(titular_sheet, 'B18', 'Sergio Andres')

            titular_sheet = modify_excel(titular_sheet, 'B20', 'Trujillo')

            titular_sheet = modify_excel(titular_sheet, 'B22', 'Garcia')

            titular_sheet = modify_excel(titular_sheet, 'L18', None)

            # Leer la hoja OPERATIVIDAD y modificar los datos
            operatividad_sheet = workbook['OPERATIVIDAD']
            operatividad_sheet = modify_excel(operatividad_sheet, 'B7', 'Sergio Garcia')

        # Guardar el archivo modificado en S3
        modified_key = 'modified_plantilla.xlsx'
        with io.BytesIO() as output_stream:
            workbook.save(output_stream)
            output_stream.seek(0)
            s3.put_object(Bucket=bucket, Key=modified_key, Body=output_stream.read())

        return {
            'statusCode': 200,
            'body': json.dumps('Archivo modificado guardado en S3')
        }

    except Exception as e:
        print(e)
        return {
            'statusCode': 500,
            'body': json.dumps('Error')
        }

def modify_excel(sheet, point, value):
    for merged_cell in sheet.merged_cells.ranges:
        if point in merged_cell:
            # Modificar la celda superior izquierda del rango combinado
            sheet[merged_cell.coord.split(':')[0]].value = value
            break
        else:
            # Si no es parte de un rango combinado, modificar la celda directamente
            sheet[point] = value
            break
    return sheet
