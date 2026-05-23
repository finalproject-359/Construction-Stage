(() => {
  const brandEl = document.querySelector('.sidebar .brand');
  if (!(brandEl instanceof HTMLElement)) return;

  brandEl.classList.add('brand-toggle');
  brandEl.tabIndex = 0;
  brandEl.setAttribute('role', 'button');
  brandEl.setAttribute('aria-haspopup', 'menu');
  brandEl.setAttribute('aria-expanded', 'false');
  brandEl.setAttribute('aria-label', 'Open account menu');

  const menu = document.createElement('div');
  menu.className = 'brand-account-menu';
  menu.setAttribute('role', 'menu');
  menu.hidden = true;
  menu.innerHTML = '<button type="button" class="brand-signout-btn" role="menuitem">Sign out</button>';

  brandEl.insertAdjacentElement('afterend', menu);

  const signOutBtn = menu.querySelector('.brand-signout-btn');

  const closeMenu = () => {
    menu.hidden = true;
    brandEl.setAttribute('aria-expanded', 'false');
  };

  const openMenu = () => {
    menu.hidden = false;
    brandEl.setAttribute('aria-expanded', 'true');
  };

  const toggleMenu = () => {
    if (menu.hidden) openMenu();
    else closeMenu();
  };

  brandEl.addEventListener('click', toggleMenu);
  brandEl.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      toggleMenu();
    }
    if (event.key === 'Escape') closeMenu();
  });

  document.addEventListener('click', (event) => {
    const target = event.target;
    if (!(target instanceof Node)) return;
    if (!brandEl.contains(target) && !menu.contains(target)) closeMenu();
  });

  signOutBtn?.addEventListener('click', () => {
    sessionStorage.removeItem('costrackAuth');
    sessionStorage.removeItem('costrackPlayDashboardIntro');
    localStorage.removeItem('costrackAuth');
    localStorage.removeItem('costrackRememberMe');
    window.location.assign('login.html');
  });
})();
