const { test, expect } = require('@playwright/test');
const fs = require('fs');

test('Application Boot, Login, and Rules Gateway Routing', async ({ page }) => {
    let rawHtml = fs.readFileSync('tests/merged.html', 'utf8');

    const scriptMock = `
      <script>
         window.google = {
             script: {
                 run: {
                     withSuccessHandler: function(cb) { this.successCb = cb; return this; },
                     withFailureHandler: function(cb) { this.failureCb = cb; return this; },
                     verifyUser: function(name, pin) {
                         setTimeout(() => {
                             if (name === 'TestUser' && pin === '1234') {
                                 this.successCb({ status: 'Success', token: 'mock-token-123' });
                             } else {
                                 this.successCb({ status: 'Invalid PIN' });
                             }
                         }, 50);
                     },
                     getRulesAndTips: function() {
                         setTimeout(() => { this.successCb({ rules: ['Rule 1', 'Rule 2'], tips: [] }); }, 50);
                     },
                     getPublicCalendarData: function() {
                         setTimeout(() => { this.successCb({ calendarData: [], weekendData: [], holidayData: [], turnQueue: [], currentRound: 1, currentPhase: 'SETUP' }); }, 50);
                     },
                     getDashboardData: function() {
                         setTimeout(() => {
                             this.successCb({
                                currentUser: { name: 'TestUser', status: 'Active', weeksSelectedCount: 0, selectedWeeks: [] },
                                currentPhase: 'VACATION_RANDOM',
                                currentRound: 2
                             });
                         }, 50);
                     },
                     submitRulesAcknowledgment: function() {
                         setTimeout(() => { this.successCb({ success: true }); }, 50);
                     },
                     processSelection: function(payload) {
                         setTimeout(() => {
                             if (payload.token === 'mock-token-123') {
                                 this.successCb({ success: true, message: 'Vacation selection successful.' });
                             } else {
                                 this.successCb({ success: false, message: 'Invalid session.' });
                             }
                         }, 50);
                     }
                 }
             }
         };
      </script>
    `;
    rawHtml = rawHtml.replace('</body>', scriptMock + '</body>');
    await page.setContent(rawHtml);

    // Login -> Rules -> Dashboard
    await page.waitForSelector('#login-form');
    await page.evaluate(() => {
        const sel = document.getElementById('login-name');
        const opt = document.createElement('option');
        opt.value = 'TestUser'; opt.text = 'TestUser';
        sel.appendChild(opt);
    });

    await page.selectOption('#login-name', 'TestUser');
    await page.fill('#login-pin', '1234');
    await page.click('#login-btn');

    await expect(page.locator('#rules-gateway')).toBeVisible();
    await page.selectOption('#volunteer-select', 'yes');
    await page.selectOption('#transfer-select', 'both');
    await page.click('#accept-btn');

    await expect(page.locator('#dashboard-card')).toBeVisible();
    await expect(page.locator('#dashboard-welcome')).toBeVisible();

    // Check Controls Visibility
    await expect(page.locator('#selection-controls')).toBeVisible();
    await expect(page.locator('#start-selection-btn')).toBeVisible();

    // Simulate Vacation Selection Click
    await page.click('#start-selection-btn');

});

test('Session Expiration UI Gracefully Reverts', async ({ page }) => {
    let rawHtml = fs.readFileSync('tests/merged.html', 'utf8');
    const scriptMock = `
      <script>
         window.google = { script: { run: {
             withSuccessHandler: function(cb) { this.successCb = cb; return this; },
             withFailureHandler: function(cb) { this.failureCb = cb; return this; },
             getPublicCalendarData: function() { this.successCb({ calendarData: [], turnQueue: [] }); },
             processSelection: function() {
                 setTimeout(() => { this.successCb({ success: false, message: 'Invalid session' }); }, 50);
             }
         }}};
      </script>
    `;
    rawHtml = rawHtml.replace('</body>', scriptMock + '</body>');
    await page.setContent(rawHtml);

    // Inject mock state explicitly
    await page.evaluate(() => {
        window.appState = {
            dashboardData: { currentPhase: 'VACATION', currentUser: { status: 'Active' } },
            currentUserName: 'TestUser',
            sessionToken: 'bad-token',
            dashboardData: { currentPhase: 'VACATION', currentUser: { status: 'Active' } }
        };
        document.getElementById('login-card').classList.add('hidden');
        document.getElementById('dashboard-card').classList.remove('hidden');
    });

    await expect(page.locator('#dashboard-card')).toBeVisible();
    // Use the native submit
    await page.evaluate(() => { window.appFunctions.submitSelection(); });

    await expect(page.locator('#toast-container')).toBeVisible();
    await expect(page.locator('#toast-container')).toContainText('Session expired');
    await expect(page.locator('#login-card')).not.toHaveClass(/hidden/);
});
