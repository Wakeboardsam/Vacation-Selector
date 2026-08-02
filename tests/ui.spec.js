const { test, expect } = require('@playwright/test');

test('Mock Rules Gateway Intercept', async ({ page }) => {
  await page.setContent(`
    <!DOCTYPE html>
    <html>
      <body>
         <div id="rules-gateway" style="display:block;">
            <h1>Rules and Tips</h1>
            <form id="rules-form">
                <select required id="volunteer-select"><option value="yes">Yes</option></select>
                <select required id="transfer-select"><option value="none">None</option></select>
                <button id="accept-btn">Accept</button>
            </form>
         </div>
         <div id="dashboard" style="display:none;">
            <h1>Dashboard</h1>
         </div>
         <script>
            document.getElementById('accept-btn').addEventListener('click', (e) => {
                e.preventDefault();
                document.getElementById('rules-gateway').style.display = 'none';
                document.getElementById('dashboard').style.display = 'block';
            });
         </script>
      </body>
    </html>
  `);

  await expect(page.locator('#rules-gateway')).toBeVisible();
  await expect(page.locator('#dashboard')).toBeHidden();

  await page.locator('#accept-btn').click();

  await expect(page.locator('#rules-gateway')).toBeHidden();
  await expect(page.locator('#dashboard')).toBeVisible();
});

test('Session Expiration UI Gracefully Reverts', async ({ page }) => {
   await page.setContent(`
    <!DOCTYPE html>
    <html>
      <body>
         <div id="login-card" style="display:none;"><h2>Login</h2></div>
         <div id="dashboard-card" style="display:block;">
            <button id="submit-btn">Submit Pick</button>
            <div id="toast" style="display:none;"></div>
         </div>
         <script>
            document.getElementById('submit-btn').addEventListener('click', () => {
               // mock expiration
               document.getElementById('toast').innerText = 'Session expired. Please log in again.';
               document.getElementById('toast').style.display = 'block';
               document.getElementById('dashboard-card').style.display = 'none';
               document.getElementById('login-card').style.display = 'block';
            });
         </script>
      </body>
    </html>
  `);

  await expect(page.locator('#dashboard-card')).toBeVisible();
  await page.locator('#submit-btn').click();
  await expect(page.locator('#toast')).toHaveText('Session expired. Please log in again.');
  await expect(page.locator('#login-card')).toBeVisible();
});

test('Weekend Warning Advisory Visibility', async ({ page }) => {
   await page.setContent(`
    <!DOCTYPE html>
    <html>
      <body>
         <div id="weekend-picker-modal" class="modal">
            <div id="weekend-picker-warnings">Advisory: Adjacent to Vacation</div>
            <button id="weekend-confirm-btn">Select</button>
         </div>
      </body>
    </html>
  `);

  await expect(page.locator('#weekend-picker-warnings')).toBeVisible();
  await expect(page.locator('#weekend-picker-warnings')).toHaveText('Advisory: Adjacent to Vacation');
  // Confirm button must remain enabled
  await expect(page.locator('#weekend-confirm-btn')).toBeEnabled();
});
