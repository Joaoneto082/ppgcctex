# latexmk (LaTeX Workshop): não vetar bibtex só porque um .bib do sistema
# não passou na verificação de existência (Windows / abntex2-options.bib).
# 2 = rodar bibtex/biber quando o .aux indicar, sem exigir "todos os .bib existem".
$bibtex_use = 2;

# Pacote glossaries: gera .gls/.acr com makeglossaries (igual ao compile.ps1)
add_cus_dep('glo', 'gls', 0, 'run_makeglossaries');
add_cus_dep('acn', 'acr', 0, 'run_makeglossaries');

sub run_makeglossaries {
  my ($base) = @_;
  return system('makeglossaries', $base);
}
