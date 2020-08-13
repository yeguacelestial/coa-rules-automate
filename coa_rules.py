import pandas as pd

"""
    TODO:
        - Verificar todos los 'Business titles' existentes antes de pasarse a Rules
        - Procesar archivo de RULES
"""

def main():
    """ Main function """
    """
        - ENTRADA:
            * archivo_coa => Archivo de Excel de COA
            * archivo_rules => Archivo de Excel de RULES

            * impact => Leer dato de Impact
            * category_code => Leer dato de category_code
            * multiplant => Leer dato de Multiplant, yes or no
            * plant_impacted => Leer dato de planta impactada
            * updating_type => Leer dato de modalidad de actualización

        - SALIDA:
            * approval_by => Dato de por quién se tiene que aprobar
            * inform_to => Dato de a quién se tiene que informar
    
    """

    try:
        # Leer archivos de entrada
        archivo_coa = str(input('[*] Nombre del archivo de COA: '))
        global coa_lista
        coa_lista = leer_excel_coa('COA.xlsx').values.tolist()

        archivo_rules = str(input('[*] Nombre del archivo de RULES: '))
        rules_dataframe = leer_excel_rules(archivo_rules)
        
        # Leer datos de entrada
        impact_value = float(input('[*] IMPACT VALUE: '))
        category_code = str(input('[*] CATEGORY CODE (COMMODITY): '))
        multiplant = str(input('[*] Multiplant? [Yes/no]: '))
        plant_impacted = str(input('[*] Plant impacted: '))
        updating_type = str(input('[*] Updating type: '))

        # Evaluar rango de impact_value
        get_business_titles = rango_impact_value(
            rules_dataframe=rules_dataframe, 
            impact_value=impact_value, 
            updating_type=updating_type)

        # Business Titles de A e I
        approve_business_title = get_business_titles['Approve']
        inform_business_title = get_business_titles['Inform']

        print(f"\n[+] La requisión del buyer de {category_code} es de ${impact_value}")
        print(f"[+] Aprobar: {approve_business_title} de {category_code}")
        print(f"[+] Informar: {inform_business_title} de {plant_impacted}")
        print(get_business_titles)

        # Buscar empleado que debe aprobar la requisición
        approve_employee = get_approve_employee(coa_lista, approve_business_title, category_code)
        print(f"[+] Aprobador de la requisicion ({approve_business_title} de {category_code}): {approve_employee}")

        # Buscar empleados a quien se les debe informar de la planta afectada
        inform_employee = get_inform_employee(coa_lista, inform_business_title, plant_impacted)
        print(f"[+] Informar de la planta afectada ({inform_business_title} de {plant_impacted}): {inform_employee}")

        # Debug
        # print(coa_lista)

    except FileNotFoundError:
        print("[-] Error: No se pudo encontrar el archivo Excel.")
        print("[*] Recuerda escribir el nombre del archivo de Excel y su extensión.")

    except ValueError:
        print("[-] Error: escribe un valor de impacto correcto.")

    except:
        print("[-] Error: Algo salió mal...")
        raise

    return


def leer_excel_coa(archivo_coa:str):
    """ Leer el archivo de Excel donde vienen los datos COA
        ENTRADA:
            - archivo_coa => nombre del archivo de Excel donde vienen las COA
        SALIDA:
            - diccionario_coa => Diccionario donde viene tabulada toda la información de COA
    """

    # Convertir la primera hoja del archivo Excel de COA a diccionario
    lista_coa = excel_a_dic(archivo_coa, 0, exportar_solo_dataframe=True)

    return lista_coa


def leer_excel_rules(archivo_rules:str):
    """ Leer el archivo de Excel donde vienen los datos de Rules 
        ENTRADA: 
            - archivo_rules => nombre del archivo de Excel donde vienen las rules
        SALIDA:
            - diccionario_rules => Diccionario de información de Rules
    """

    # Convertir la primera hoja del archivo Excel de RULES a diccionario
    dataframe_rules = excel_a_dic(archivo_rules, 0, exportar_solo_dataframe=True)
    
    return dataframe_rules


def excel_a_dic(nombre_archivo:str, hoja:int, exportar_solo_dataframe:bool=False):
    """ Convertir la hoja de un archivo Excel a un diccionario
        ENTRADA:
            - archivo => Nombre del archivo Excel donde se encuentra la hoja
            - hoja => Número de hoja a automatizar (0 es la primera y se cuenta de forma ascendente,
              -1 es la última hoja del archivo)
        SALIDA:
            - hoja_a_dic => Diccionario de la hoja de Excel   
    """

    # Leer archivo Excel
    archivo_excel = pd.ExcelFile(nombre_archivo)

    # Leer hoja, y convertirlo a un Panda Dataframe
    hoja_dataframe = archivo_excel.parse(archivo_excel.sheet_names[hoja])

    # Si solo se quiere exportar el dataframe...
    if exportar_solo_dataframe:
        return hoja_dataframe

    # De lo contrario...
    else:
        # Convertir hoja a diccionario
        hoja_a_dic = hoja_dataframe.to_dict()
        return hoja_a_dic


def rango_impact_value(rules_dataframe, impact_value:float, updating_type:str):
    """ EVALUAR EL RANGO DE impact_value
        ENTRADA:
            - rules_dataframe => Objeto tipo Pandas Dataframe de la hoja de rules
            - impact_value => Valor numérico de entrada tipo float, introducido por el usuario
            - updating_type => Tipo de actualización, introducido por el usuario ('Negotation Events' o 'Price Change')
        
        SALIDA:
            - business_titles => Valor string de columna donde se encuentren letras A e I
    """
    # From $-5K to $5K
    if (impact_value > -5000) and (impact_value < 5000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=1)

    # Greater than $5K up to $10K or less than $-5K up to $-10K
    elif (impact_value > 5000 and impact_value < 10000) or (impact_value < -5000 and impact_value > -10000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=2)

    # Greater than 10K up to $50K or less than -S10K up to $-50K
    elif (impact_value > 10000 and impact_value < 50000) or (impact_value < -10000 and impact_value > -50000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=3)

    # Greater than $50K up to $100K or less than $-50K up to $-100K
    elif (impact_value > 10000 and impact_value < 50000) or (impact_value < -10000 and impact_value > -50000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=4)

    # Greater than $100k up to $300a or less than $-100k up to $-300K
    elif (impact_value > 10000 and impact_value < 50000) or (impact_value < -10000 and impact_value > -50000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=5)

    # Greather than $300K or Less than -$300K
    elif (impact_value > 10000 and impact_value < 50000) or (impact_value < -10000 and impact_value > -50000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=6)
    
    return business_titles


def get_business_titles(rules_dataframe, updating_type, rules_excel_row):
    ''' EVALUAR BUSINESS TITLES DEPENDIENDO DEL RANGO DE IMPACTO
        ENTRADA:
            - rules_dataframe => Objeto Pandas Dataframe de Rules
            - updating_type => Tipo de updating_type (Negotiation Events o Price Change)
            - rules_excel_row => Renglón de Excel donde se encuentra el rango del impact_value
        
        SALIDA:
            - roles_dict => Valor string de columna donde se encuentren las letras A e I, y el Business Title asignado.
    '''
    coa_available_business_titles = [row[1] for row in coa_lista]

    # Seccion de Negotiation Events
    if updating_type == 'Negotiation Events':

        rules_rango_renglon = rules_dataframe.iloc[rules_excel_row][1:].to_dict()

        roles_dict = filtrar_business_titles(rules_rango_renglon, coa_available_business_titles)
        return roles_dict

    # Seccion de Price Change
    elif updating_type == 'Price Change':

        rules_rango_renglon = rules_dataframe.iloc[rules_excel_row+7][1:].to_dict()

        roles_dict = filtrar_business_titles(rules_rango_renglon, coa_available_business_titles)
        return roles_dict
    
    else:
        print("[-] Error: Updating type inválido.")
        print("[*] Updating types disponibles: 'Price Change' y 'Negotiation Events'")


def filtrar_business_titles(rules_rango_renglon, coa_available_business_titles):
    roles_dict = {}

    for business_title, rol in rules_rango_renglon.items():

        if type(0.0) == type(rol):
            pass

        # Validar que el rol existe en los business_titles de COA
        elif business_title in coa_available_business_titles:
            print(f"[+] Business Title {business_title} SE ENCUENTRA EN COA")

            if 'A' in rol:
                roles_dict['Approve'] = business_title

            if 'I' in rol:
                roles_dict['Inform'] = business_title
            
            if 'C' in rol:
                roles_dict['Consult'] = business_title

        # Si no existe el business_title, seguir iterando
        else:
            pass

    return roles_dict


def get_approve_employee(coa_list:list, approve_business_title:str, category_code:str):
    """ BUSCAR EMPLEADO QUE DEBE APROBAR LA REQUISICIÓN
        ENTRADA:
            - coa_list => Diccionario de la hoja de Excel de COA
            - approve_business_title => BT de la persona que debe aprobar la requisición
            - category_code => Valor de "Commodity" de la persona que debe aprobar la requisición
        SALIDA:
            - approve_employee => Nombre del empleado que debe aprobar la requisición
    """
    for employee in coa_list:
        # Buscar empleado con el Business Title y el Category Code adecuados
        if employee[1] == approve_business_title and category_code in employee[2]:
            print(f"[*] EMPLEADO {approve_business_title} DE {category_code} => {employee[0]}")
            # Regresar el nombre del empleado
            return employee[0]


def get_inform_employee(coa_list:list, inform_business_title:str, plant_impacted:str):
    """ BUSCAR EMPLEADOS A QUIENES SE LES DEBE INFORMAR DE LAS PLANTAS AFECTADAS
        ENTRADA:
            - coa_list => Lista de la hoja de Excel de COA
            - inform_business_title => BT de las personas a quien se les debe informar
            - plant_impacted => Plantas impactadas por la requisición
        
        SALIDA:
            - inform_employee => Empleado/s a quien/es se le/s debe informar de la/s planta/s afectada/s
    """
    for employee in coa_list:
        if employee[1] == inform_business_title and employee[3] == plant_impacted:
            return employee[0]

# TODO
def get_consult_employee(coa_diccionario:dict, consult_business_title:str, commodity:str, plant_impacted:str):
    """ OBTENER EMPLEADO CON LA LETRA C (CONSULTED)
        ENTRADA:
            - coa_diccionario => Diccionario de COA
            - consult_business_title => String 'Business Title' del empleado
            - commodity => String de 'Commodity' del empleado
            - plant_impacted => String de 'Planta impactada' del empleado

        SALIDA:
            - consult_employee => String de 'Employee name' asociado a 'Commodity' o 'Plant impacted'
    """

    # Primero, se intenta buscar al 'Employee name' con el business title asociado al commodity.
    # for k,v in coa_diccionario['Business Title'].items():


    # Si no se encuentra el 'Employee name' asoaicado con el 'Commodity'...
    #  se buscará al 'Employee name' con 'Plant impacted asociado'


    return


if __name__ == '__main__':
    main()
    pass