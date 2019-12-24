UPDATE admCategorias 
	SET admCategorias.Descricao01 = REPLACE(admCategorias.Descricao01, 'AZS', '')
		WHERE admCategorias.Descricao01 LIKE '%AZS%' and admCategorias.codRelacao in ('1346');

SELECT *
FROM admCategorias
		WHERE admCategorias.Descricao01 LIKE '%AZS%' and admCategorias.codRelacao in ('1346');


SELECT *
FROM admCategorias
		WHERE admCategorias.codRelacao in ('1346');
