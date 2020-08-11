import pandas as pd

# FEATURE DESCRIPTION:
"""
    - HOJA DE RULES: Tomar datos de entrada y evaluar en qué rango se encuentra impact_value
        impact_value -> 40,000
        impact_value_rango -> Renglón 3
                            => Renglon de rango de impact_value
        
        category_code -> K4
                        => Número de renglón en hoja_de_COA[columna_commodity][category_code]
        
        multiplant -> Yes
                    => Elegir si se afectará más de una planta
        
        plant_impacted -> Tijuana/Tlaxcala
                        => Nombre de la o las plantas
        
    - HOJA DE RULES: Después de evaluar el rango, buscar quien tiene los valores A e I
        approved_columna -> RCM
                            => hoja_de_rules[impact_value_rango][nombreColumna_de_A]

        informed_columna -> PPM
                            => hoja_de_rules[impact_value_rango][nombreColumna_de_I]
                            
    - HOJA DE COA: Sacar quién aprobará la requisición.
                    Se necesita sacar el 'Employee name' donde
                        Columna 'Business title' => approved_columna
                        Columna 'Commodity' => category_code

        approved_employee_name -> Aram Gonzalez
        approved_employee_name => IF (Columna 'Business Title' == approved_columna) 
                                    AND (Columna 'Commodity' == category_code):
                                        approved_employee_name = hoja_de_COA[renglón de 'Business title' y 'Commodity'][Columna 'Employee name']
    
    - HOJA DE COA: Sacar a quien se informará de las plantas afectadas (HACER ESTO EN CADA PLANTA INDICADA, SI SON MAS DE UNA)
        informed_employee_name -> 
        informed_employee_name => IF (Columna 'Business Title' == informed_columna)
                                    AND (Columna 'Plant' == plant_impacted):
                                        informed_employee_name = hoja_de_COA[renglón de 'Business title' y 'Commodity'][Columna 'Employee name']

    - Imprimir valores 'aproved_employee_name' e 'informed_employee_name'
"""

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
        global coa_diccionario
        coa_diccionario = leer_excel_coa(archivo_coa)
        archivo_rules = str(input('[*] Nombre del archivo de RULES: '))
        rules_dataframe = leer_excel_rules(archivo_rules)
        
        # Leer datos de entrada
        impact_value = float(input('[*] IMPACT VALUE: '))
        category_code = str(input('[*] CATEGORY CODE (COMMODITY): '))
        multiplant = str(input('[*] Multiplant? [Yes/no]: '))
        plant_impacted = str(input('[*] Plant impacted: '))
        updating_type = str(input('[*] Updating type: '))
        
        # Business Titles disponibles (diccionario)
        coa_available_business_titles = coa_diccionario['Business Title']

        # Evaluar rango de impact_value
        """
            impact_value => Renglon de rango de impact_value
        """
        get_business_titles = rango_impact_value(
            rules_dataframe=rules_dataframe, 
            impact_value=impact_value, 
            updating_type=updating_type)
        print(get_business_titles)


    except FileNotFoundError:
        print("[-] Error: No se pudo encontrar el archivo Excel.")
        print("[*] Recuerda escribir el nombre del archivo de Excel y su extensión.")

    except ValueError:
        print("[-] Error: escribe un valor de impacto correcto.")

    except:
        print("[-] Error: Algo salió mal. Inténtalo de nuevo.")
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
    diccionario_coa = excel_a_dic(archivo_coa, 0)

    return diccionario_coa


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
    if (impact_value > 5000) and (impact_value < 5000):
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
    ''' EVALUAR COMMODITY DEPENDIENDO DEL RANGO DE IMPACTO
        ENTRADA:
            - rules_dataframe => Objeto Pandas Dataframe de Rules
            - updating_type => Tipo de updating_type (Negotiation Events o Price Change)
            - rules_excel_row => Renglón de Excel donde se encuentra el rango del impact_value
        
        SALIDA:
            - approved_value, informed_value => Valor string de columna donde se encuentren las letras A e I
    '''
    coa_available_business_titles = coa_diccionario['Business Title']
    roles_dict = {}

    # Seccion de Negotiation Events
    if updating_type == 'Negotiation Events':

        rules_rango_renglon = rules_dataframe.iloc[rules_excel_row][1:].to_dict()

        # TODO: Solo verificar los bussiness_titles existentes en el archivo de COA.xlsx
        for business_title, rol in rules_rango_renglon.items():

            # Validar que el rol existe en los business_titles de COA
            if business_title in coa_available_business_titles.values():
                print(f"[+] {business_title} SE ENCUENTRA EN COA")
                if type(0.0) == type(rol):
                    pass

                if 'A' in rol:
                    print(f'BUSINESS TITLE: {business_title} => ROL: {rol}')
                    roles_dict['Approve'] = business_title

                if 'I' in rol:
                    print(f'BUSINESS TITLE: {business_title} => ROL: {rol}')
                    roles_dict['Inform'] = business_title

            # Si no existe el business_title, seguir iterando
            else:
                print(f"[-] El ROL {rol} no está en {coa_available_business_titles.values()}")
                pass

        return roles_dict

    # Seccion de Price Change
    elif updating_type == 'Price Change':

        # Seccion de Price Change
        rules_rango_renglon = rules_dataframe.iloc[rules_excel_row+7][1:].to_dict()

        for business_title, rol in rules_rango_renglon.items():
            # Validar que el rol existe en los business_titles de COA
            if business_title in coa_available_business_titles.values():
                print(f"[+] {business_title} SE ENCUENTRA EN COA")
                if type(0.0) == type(rol):
                    pass

                if 'A' in rol:
                    print(f'BUSINESS TITLE: {business_title} => ROL: {rol}')
                    roles_dict['Approve'] = business_title

                if 'I' in rol:
                    print(f'BUSINESS TITLE: {business_title} => ROL: {rol}')
                    roles_dict['Inform'] = business_title

            # De lo contrario, seguir iterando
            else:
                print(f"[-] El BT {business_title} no está en el archivo de COA.")
                pass
        
        return roles_dict


if __name__ == '__main__':
    main()
    # pass