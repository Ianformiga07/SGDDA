// JavaScript Document// JavaScript Document
//adiciona mascara de cnpj

function MascaraCNPJ(cnpj)
{	
   if(mascaraInteiro(cnpj)==false)
   {		
     event.returnValue = false;	
   }		
    return formataCampo(cnpj, '00.000.000/0000-00', event);
}
//adiciona mascara de cep
function MascaraCep(cep)
{		
  if(mascaraInteiro(cep)==false)
  {		
   event.returnValue = false;	
  }		
   return formataCampo(cep, '00.000-000', event);
}
//adiciona mascara de data
function MascaraData(data)
{	
  if(mascaraInteiro(data)==false)
  {		
    event.returnValue = false;	
  }		
   return formataCampo(data, '00/00/0000', event);
}
//adiciona mascara ao telefone
function MascaraTelefone(tel)
{		
   if(mascaraInteiro(tel)==false)
   {		
     event.returnValue = false;	
   }		
   return formataCampo(tel, '(00) 0000-0000', event);
}
//adiciona mascara celular
function MascaraCelular(tel)
{		
   if(mascaraInteiro(tel)==false)
   {		
     event.returnValue = false;	
   }		
   return formataCampo(tel, '(00) 00000-0000', event);
}
//adiciona mascara ao CPF
function MascaraCPF(cpf)
{	
  if(mascaraInteiro(cpf)==false)
  {		
    event.returnValue = false;	
  }		
  return formataCampo(cpf, '000.000.000-00', event);
}
//valida telefone
function ValidaTelefone(tel)
{	
  exp = /\(\d{2}\)\ \d{4}\-\d{4}/	
  if(!exp.test(tel))
  {
	 alert("Telefone Invalido! (00) 0000-0000")		
 	 return false;
  }
  return true;
}
//valida celular
function ValidaCelular(tel)
{	
  exp = /\(\d{2}\)\ \d{5}\-\d{4}/	
  if(!exp.test(tel))
  {
	 alert("Celular Invalido! (00) 00000-0000")	
 	 return false;
  }
  return true;
	 
}
//valida CEP
function ValidaCep(cep)
{	
   exp = /\d{2}\.\d{3}\-\d{3}/	
   if(!exp.test(cep))
   {
	 alert("Cep Invalido! 00.000-000")
     return false;
   }
   return true;
}


function ValidaData(data)
	{
		
		
		day = data.substring(0,2);
        month = data.substring(3,5);
        year = data.substring(6,10);
		var daysInMonth = makeArray(12);
		daysInMonth[1] =  31;	//Janeiro
		daysInMonth[2] =  29;	// daysInFebruary
		daysInMonth[3] =  31;	//Março
		daysInMonth[4] =  30;	//Abril
		daysInMonth[5] =  31;	//Maio
		daysInMonth[6] =  30;	//Junho
		daysInMonth[7] =  31;	//Julho
		daysInMonth[8] =  31;	//Agosto
		daysInMonth[9] =  30;	//Setembro
		daysInMonth[10] = 31;	//Outubro	
		daysInMonth[11] = 30;	//Novembro
		daysInMonth[12] = 31;	//Dezembro
		// verifica se os dias, meses e anos sao inteiros válidos
		if ( isNaN(year) || (year.indexOf(".")!=-1) || (year.length==0) || (year.length!=4) || (Number(year)<0) )
			return (false);
		if ( isNaN(month) || (month.indexOf(".")!=-1) || (month.length == 0) || (Number(month)<1) || (Number(month)>12) )
			return (false);
		if ( isNaN(day) || (day.indexOf(".")!=-1) || (day.length==0) || (Number(day)<1))
			return (false);
		var intYear = parseInt(year);
		var intMonth = parseFloat(month);
		var intDay = parseInt(day);
		// verifica dias inválidos, exceto para fevereiro
		if (intDay > (daysInMonth[intMonth]))
		   return false; 
		if ( (intMonth == 2) && (intDay > daysInFebruary(intYear)) )
			return false;  
	  return true;
	}
	
function makeArray(n)
{
   for (var i = 1; i <= n; i++)
	this[i] = 0;
   return this;
}

function daysInFebruary (year)
{
   return (  ((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0) ) ) ? 29 : 28 );
}



//valida o CPF digitado
function ValidarCPF(cpf)
{	

  
    cpf = cpf.replace(/[^\d]+/g,'');    
    if(cpf == '') return false; 
    // Elimina CPFs invalidos conhecidos    
    if (cpf.length != 11 || 
        cpf == "00000000000" || 
        cpf == "11111111111" || 
        cpf == "22222222222" || 
        cpf == "33333333333" || 
        cpf == "44444444444" || 
        cpf == "55555555555" || 
        cpf == "66666666666" || 
        cpf == "77777777777" || 
        cpf == "88888888888" || 
        cpf == "99999999999")
            return false;       
    // Valida 1o digito 
    add = 0;    
    for (i=0; i < 9; i ++)       
        add += parseInt(cpf.charAt(i)) * (10 - i);  
        rev = 11 - (add % 11);  
        if (rev == 10 || rev == 11)     
            rev = 0;    
        if (rev != parseInt(cpf.charAt(9)))     
            return false;       
    // Valida 2o digito 
    add = 0;    
    for (i = 0; i < 10; i ++)        
        add += parseInt(cpf.charAt(i)) * (11 - i);  
    rev = 11 - (add % 11);  
    if (rev == 10 || rev == 11) 
        rev = 0;    
    if (rev != parseInt(cpf.charAt(10)))
	    return false;       
	return true;   
}
//valida numero inteiro com mascara
function mascaraInteiro()
{	
  if (event.keyCode < 48 || event.keyCode > 57)
    {		
      event.returnValue = false;		
      return false;	
     }	
     return true;
}
//valida o CNPJ digitado


function validarCNPJ(cnpj) {
 
    cnpj = cnpj.replace(/[^\d]+/g,'');
 
    if(cnpj == '') return false;
     
    if (cnpj.length != 14)
        return false;
 
    // Elimina CNPJs invalidos conhecidos
    if (cnpj == "00000000000000" || 
        cnpj == "11111111111111" || 
        cnpj == "22222222222222" || 
        cnpj == "33333333333333" || 
        cnpj == "44444444444444" || 
        cnpj == "55555555555555" || 
        cnpj == "66666666666666" || 
        cnpj == "77777777777777" || 
        cnpj == "88888888888888" || 
        cnpj == "99999999999999")
        return false;
         
    // Valida DVs
    tamanho = cnpj.length - 2
    numeros = cnpj.substring(0,tamanho);
    digitos = cnpj.substring(tamanho);
    soma = 0;
    pos = tamanho - 7;
    for (i = tamanho; i >= 1; i--) {
      soma += numeros.charAt(tamanho - i) * pos--;
      if (pos < 2)
            pos = 9;
    }
    resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    if (resultado != digitos.charAt(0))
        return false;
         
    tamanho = tamanho + 1;
    numeros = cnpj.substring(0,tamanho);
    soma = 0;
    pos = tamanho - 7;
    for (i = tamanho; i >= 1; i--) {
      soma += numeros.charAt(tamanho - i) * pos--;
      if (pos < 2)
            pos = 9;
    }
    resultado = soma % 11 < 2 ? 0 : 11 - soma % 11;
    if (resultado != digitos.charAt(1))
         return false;
           
    return true;
    
}

//formata de forma generica os campos
function formataCampo(campo, Mascara, evento) 
{ 	
   var boleanoMascara; 		
   var Digitato = evento.keyCode;	
   exp = /\-|\.|\/|\(|\)| /g	
   campoSoNumeros = campo.value.toString().replace( exp, "" );    	
   var posicaoCampo = 0;	 	
   var NovoValorCampo="";	
   var TamanhoMascara = campoSoNumeros.length;; 		
   if (Digitato != 8) 
    { // backspace 		
       for(i=0; i<= TamanhoMascara; i++) 
         { 			
            boleanoMascara  = ((Mascara.charAt(i) == "-") || (Mascara.charAt(i) == ".")	|| (Mascara.charAt(i) == "/"))
            boleanoMascara  = boleanoMascara || ((Mascara.charAt(i) == "(") || (Mascara.charAt(i) == ")") || (Mascara.charAt(i) == " ")) 			
            if (boleanoMascara) 
            { 				
               NovoValorCampo += Mascara.charAt(i); 				  
               TamanhoMascara++;			
             }
             else 
                 { 				
                   NovoValorCampo += campoSoNumeros.charAt(posicaoCampo); 				
                   posicaoCampo++; 			  }	   	 		  
                  }	 		
                  campo.value = NovoValorCampo;		  
                  return true; 	
                  }
               else { 		
                 return true; 	
       }
}

function maiuscula(z){
        v = z.value.toUpperCase();
        z.value = v;
    }
	
function somente_numero(campo){   
var digits="0123456789"   
var campo_temp   
    for (var i=0;i<campo.value.length;i++){   
        campo_temp=campo.value.substring(i,i+1)   
        if (digits.indexOf(campo_temp)==-1){   
            campo.value = campo.value.substring(0,i);   
        }   
    }   
} 

function CPFCNPJ(CPFCNPJ)
{
	
	if (CPFCNPJ.value.length < 14)
	 MascaraCPF(CPFCNPJ);
	else
	 MascaraCNPJ(CPFCNPJ);
}