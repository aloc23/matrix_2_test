document.addEventListener('DOMContentLoaded', function() {
  const investmentHub = document.getElementById('investment-hub');
  const projectHub = document.getElementById('project-hub');
  const btnInvestment = document.getElementById('show-investment-hub');
  const btnProject = document.getElementById('show-project-hub');

  btnInvestment.addEventListener('click', function() {
    investmentHub.style.display = '';
    projectHub.style.display = 'none';
    btnInvestment.classList.add('active');
    btnProject.classList.remove('active');
  });
  btnProject.addEventListener('click', function() {
    investmentHub.style.display = 'none';
    projectHub.style.display = '';
    btnProject.classList.add('active');
    btnInvestment.classList.remove('active');
  });
});