"""
    TODO
        - Validate PSL and PPM business titles on get_approve_employee and get_inform_employee
"""

import pandas as pd


def main():
    """ Main function """
    """
        - ENTRADA:
            * archivo_coa => Archivo de Excel de COA
            * archivo_rules => Archivo de Excel de RULES

            * impact => Leer dato de Impact
            * category_code => Leer dato de category_code
            * multiplant => Leer dato de Multiplant, yes or no
            * plant_impacted => Leer dato de plantas impactadas
            * updating_type => Leer dato de modalidad de actualización

        - SALIDA:
            * approval_by => Dato de por quién se tiene que aprobar
            * inform_to => Dato de a quién se tiene que informar
            * consult_to => Dato de a quién se tiene que consultar
    
    """

    try:
        """ENTRADA"""
        # Leer archivos de entrada
        archivo_coa = str(input('[*] Nombre del archivo de COA: '))
        global coa_lista
        coa_lista = leer_excel_coa(archivo_coa).values.tolist()

        archivo_rules = str(input('[*] Nombre del archivo de RULES: '))
        rules_dataframe = leer_excel_rules(archivo_rules)
        
        # Leer datos de entrada
        impact_value = float(input('[*] IMPACT VALUE (DLLS): $'))
        category_code = str(input('[*] CATEGORY CODE (COMMODITY): '))
        plant_impacted = str(input('[*] Plant or plants impacted (Si son mas de una, separarlas por coma): '))
        plant_impacted = plant_impacted.replace(' ', '')
        plant_impacted = plant_impacted.split(',')
        updating_type = str(input('[*] Updating type: '))

        # Evaluar rango de impact_value
        get_business_titles = rango_impact_value(
            rules_dataframe=rules_dataframe, 
            impact_value=impact_value, 
            updating_type=updating_type)

        # Business Titles de A e I
        approve_business_title = get_business_titles['Approve']
        inform_business_title = get_business_titles['Inform']

        """OUTPUT"""
        print(f"\n[+] La requisión del buyer de {category_code} es de ${impact_value}")

        # print("\n\n[++] BUSINESS TITLES")
        # print(f"[+] Aprobar: {approve_business_title} de {category_code}")

        # for plant in plant_impacted:
        #     print(f"\n[+] Planta {plant}")
        #     print(f"[+] Informar: {inform_business_title} de {plant}")

        # if "Consult" in get_business_titles.keys():
        #     consult_business_title = get_business_titles['Consult']

        #     for plant in plant_impacted:
        #         print(f"[+] Consultar: {consult_business_title} de {category_code}/{plant}")
        
        print("\n\n[++] EMPLEADOS")
        # Buscar empleados que deben aprobar la requisición
        approve_employees = get_approve_employee(coa_lista, approve_business_title, category_code)
        if type(approve_employees) == type(""): approve_separator = ''
        else: approve_separator = ", "
        print(f"[+] Aprobar ({approve_business_title}/{category_code}): {approve_separator.join(approve_employees)}")

        # Buscar empleados a quien se les debe informar de la planta afectada
        for plant in plant_impacted:
            print(f"\n[+] Planta {plant}")
            inform_employees = get_inform_employee(coa_lista, inform_business_title, plant)
            if type(inform_employees) == type(""): inform_separator = ''
            else: inform_separator = ", "
            print(f"[+] Informar({inform_business_title}/{plant}): {inform_separator.join(inform_employees)}")

        # Buscar empleados a quienes de les debe consultar
        for plant in plant_impacted:
            print(f"\n[+] Planta {plant}")
            consult_business_title = get_business_titles['Consult']
            consult_employees = get_consult_employee(coa_lista, consult_business_title, category_code, plant)
            if type(consult_employees) == type(""): consult_separator = ''
            else: consult_separator = ", "
            print(f"[+] Consultar({category_code}/{plant}): {consult_separator.join(consult_employees)}")


    except FileNotFoundError:
        print("[-] Error: No se pudo encontrar el archivo Excel.")
        print("[*] Recuerda escribir el nombre del archivo de Excel y su extensión.")

    except ValueError:
        print("[-] Error: escribe un valor de impacto correcto.")

    except:
        print("[-] Error: Algo salió mal...")
        raise
        print("[*] Asegúrate de escribir los datos de entrada correctamente.")
        print("\n[*] Pulsa Enter para salir del programa.")

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
    if (impact_value >= -5000) and (impact_value <= 5000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=1)

    # Greater than $5K up to $10K or less than $-5K up to $-10K
    elif (impact_value >= 5000 and impact_value <= 10000) or (impact_value <= -5000 and impact_value >= -10000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=2)

    # Greater than 10K up to $50K or less than -S10K up to $-50K
    elif (impact_value >= 10000 and impact_value <= 50000) or (impact_value <= -10000 and impact_value >= -50000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=3)

    # Greater than $50K up to $100K or less than $-50K up to $-100K
    elif (impact_value >= 50000 and impact_value <= 100000) or (impact_value <= -50000 and impact_value >= -100000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=4)

    # Greater than $100k up to $300k or less than $-100k up to $-300K
    elif (impact_value >= 100000 and impact_value <= 300000) or (impact_value <= -100000 and impact_value >= -300000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=5)

    # Greather than $300K or Less than -$300K
    elif (impact_value >= 300000) or (impact_value <= -300000):
        business_titles = get_business_titles(rules_dataframe, updating_type, rules_excel_row=6)
    
    # Any other range
    else:
        print("[-] Error: 'IMPACT VALUE' inválido.")
        exit()

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
    coa_available_business_titles = []
    for row in coa_lista:
        if row[1] not in coa_available_business_titles:
            coa_available_business_titles.append(row[1].replace('\n', ' '))

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
    roles_dict = {
        'Approve': [],
        'Inform': [],
        'Consult': [],
    }

    for business_title, rol in rules_rango_renglon.items():

        if type(0.0) == type(rol):
            pass

        # Validar que el rol existe en los business_titles de COA
        elif business_title in coa_available_business_titles:
            business_title = business_title.replace('\n', ' ')

            if 'A' in rol:
                roles_dict['Approve'].append(business_title)

            if 'I' in rol:
                roles_dict['Inform'].append(business_title)
            
            if 'C' in rol:
                roles_dict['Consult'].append(business_title)

        # Si no existe el business_title, seguir iterando
        else:
            pass

    return roles_dict


def get_approve_employee(coa_list:list, approve_business_title:list, category_code:str):
    """ BUSCAR EMPLEADO QUE DEBE APROBAR LA REQUISICIÓN
        ENTRADA:
            - coa_list => Diccionario de la hoja de Excel de COA
            - approve_business_title => Lista de BTs de las personas que deben aprobar la requisición
            - category_code => Valor de "Commodity" de la persona que debe aprobar la requisición
        SALIDA:
            - approve_employee => Nombre del empleado que debe aprobar la requisición
    """
    # Ignorar empleados que no tienen asignado un commodity
    coa_list = [employee for employee in coa_list if type(employee[3]) != type(0.0)]

    approve_employees = []
    for employee in coa_list:
        # Buscar empleado con el Business Title y el Commodity adecuados
        if employee[1] in approve_business_title and category_code in employee[3]:

            # Agregar nombre de empleados que deben aprobar la requisicion
            approve_employees.append(employee[0])
    
    # Si la lista de empleados no está vacía...
    if len(approve_employees) != 0:
        return approve_employees

    else:
        return "Empleado/s no encontrado/s"


def get_inform_employee(coa_list:list, inform_business_title:list, plant_impacted:str):
    """ BUSCAR EMPLEADOS A QUIENES SE LES DEBE INFORMAR DE LAS PLANTAS AFECTADAS
        ENTRADA:
            - coa_list => Lista de la hoja de Excel de COA
            - inform_business_title => Lista de BT de las personas a quien se les debe informar
            - plant_impacted => Plantas impactadas por la requisición
        
        SALIDA:
            - inform_employee => Empleado/s a quien/es se le/s debe informar de la/s planta/s afectada/s
    """
    # Ignorar los empleados que no tienen asignado valor de 'Plant impacted'
    coa_list = [employee for employee in coa_list if type(employee[4]) != type(0.0)]

    inform_employees = []
    for employee in coa_list:
        if employee[1] in inform_business_title and employee[4] == plant_impacted:
            inform_employees.append(employee[0])
    
    if len(inform_employees) != 0:
        return inform_employees
    
    else:
        return "Empleado/s no encontrado/s"


def get_consult_employee(coa_list:list, consult_business_title:list, commodity:str, plant_impacted:str):
    """ OBTENER EMPLEADO CON LA LETRA C (CONSULTED)
        ENTRADA:
            - coa_list => Lista de la hoja de Excel de COA
            - consult_business_title => Lista de 'Business Titles' de los empleados
            - commodity => String de 'Commodity' del empleado
            - plant_impacted => String de 'Planta impactada' del empleado

        SALIDA:
            - consult_employees => Lista de empleados asociados a 'Commodity' o 'Plant impacted'
    """
    consult_employees = []

    # Si el Business Title es PSL o PMM, se considerará a la planta
    if 'PSL' in consult_business_title or 'PPM' in consult_business_title:

        # Ignorar empleados que no tienen asignado una planta
        coa_list = [employee for employee in coa_list if type(employee[4]) != type(0.0)]

        for employee in coa_list:
            if employee[1] in consult_business_title and plant_impacted in employee[4]:
                consult_employees.append(employee[0])

    # Si el Business Title es otro, se considerará Commodity
    else:
        # Ignorar empleados que no tienen asignado un Commodity
        coa_list = [employee for employee in coa_list if type(employee[3]) != type(0.0)]

        for employee in coa_list:
            # Buscar empleado con el Business Title y el Commodity adecuados, y agregarlos a la lista de empleados
            if employee[1] in consult_business_title and commodity in employee[3]:
                consult_employees.append(employee[0])

    if len(consult_employees) == 0:
        return "Empleado/s no encontrado/s"
    
    return consult_employees


if __name__ == '__main__':
    main()
    input("\n[*] Presiona Enter para salir del programa...")