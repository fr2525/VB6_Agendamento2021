/* '***************************************************************************************
'***************************   CRIA A TABELA DE ANIMAIS   ******************************
'***************************************************************************************
'  
*/                            
CREATE TABLE tab_animais  (ID integer identity NOT NULL
						,Id_Cli integer NOT NULL
						,Nome  character(50) NOT NULL
						,Tipo_ani Int not null
						,dt_nasc date
						,pedigree CHAR(1)						
						,observacoes varchar(200)
						,cuidados_especiais varchar(100)
						,foto varchar(100)
						,dt_ult_visita date
						,operador character(10)
						,dt_Atualiza datetime, primary key (ID) )
/*
'****************************************************************************************
'****************   CRIA A TABELA DE TIPOS DE ANIMAL (CÃO/GATO/COELHO)  *****************
'****************************************************************************************
' 
*/                             
CREATE TABLE tab_tipos_an (ID integer identity NOT NULL
						,Descricao varchar(50) NOT NULL
						,operador char(10)
						,dt_Atualiza datetime
						,primary key (ID) )

/*
'****************************************************************************************
'*****************  CRIA A TABELA DE SERVICOS - BANHO/TOSA/VACINA/ETC   *****************
'****************************************************************************************
' */                              

CREATE TABLE tab_servicos (ID integer IDENTITY NOT NULL
					    ,Descricao character(50) NOT NULL
						,valor NUMERIC(12,2)
						,TEMPO_EST NUMERIC(12,2)
						,operador character(10)
						,dt_Atualiza datetime
						,primary key (ID) )
/*
'****************************************************************************************
'********************    CRIA A TABELA DE ATENDIMENTOS    *******************************
'****************************************************************************************
' 
*/
                             
CREATE TABLE tab_atendimentos (CREATE TABLE tab_atendimentos (Dt_atend timestamp NOT NULL
									, IdAnimal integer NOT NULL
									, Tipo_Atend integer NOT NULL
									, valor NUMERIC(12,2)
									, hora_saida CHAR(5)
									, operador char(10)
									, dt_Atualiza timestamp
									, primary key (dt_atend) )
        
/*
'Não tem auto incremento porque o campo chave é TIMESTAMP

'****************************************************************************************
'********************       CRIA A TABELA DE VACINAS      *******************************
'****************************************************************************************
' 
*/                             
CREATE TABLE tab_vacinas (ID integer identity not null 
						,IdAnimal integer NOT NULL
						,Dt_atend datetime NOT NULL
						,Descricao VARCHAR(100) NOT NULL
						,Valor NUMERIC(12,2)
						,DT_PROXIMA DATE
						,operador character(10)
						,dt_Atualiza datetime
						,primary key (ID)  )


CREATE TABLE tab_promocoes (ID integer identity NOT NULL 
							,Dt_inicio datetime NOT NULL
							,Dt_fim datetime NOT NULL
							,IdAnimal integer
							,IdTipoAten integer
							,Descricao VARCHAR(100) NOT NULL
							,Valor NUMERIC(12,2)
							,porcent NUMERIC(2,2)
							,operador character(10)
							,Dt_Atualiza datetime
							,primary key (ID)  )

