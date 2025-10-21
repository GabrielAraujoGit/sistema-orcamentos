|*******************************************************************************
|* txrec2400m007  B61O_a_ext
|* Imprimir confirmacao de Estoque
|* 107457
|* 2025-10-21 [10:19]
|*******************************************************************************
|* Script Type: 123
|*******************************************************************************

|*************************** declaration section *******************************
declaration:

	table	tbtrec200 |* Recebimento

	extern	domain	btorno	fire.f	fixed
	extern	domain	btorno	fire.t	fixed
	extern	domain	btdocn	docn.f
	extern	domain	btdocn	docn.t
	
	
	extern	domain	btyesno	obse.txt
	
|****************************** program section ********************************


|****************************** group section **********************************

group.1:
init.group:
	get.screen.defaults()

|****************************** choice section ********************************

choice.cont.process:
on.choice:
	execute(print.data)

choice.print.data:
on.choice:
	if rprt_open() then
		read.main.table()
		rprt_close()
	else
		choice.again()
	endif


|****************************** field section *********************************

field.fire.f:
when.field.changes:
	fire.t = fire.f

field.docn.f:
when.field.changes:
	docn.t = docn.f


|****************************** function section ******************************

functions:

function read.main.table()
{
 	select	btrec200.*
	from	btrec200
	where	btrec200._index1 inrange {:fire.f}
		                     and {:fire.t}
	order by btrec200._index1
	selectdo
		rprt_send()
	endselect
	
	
}
