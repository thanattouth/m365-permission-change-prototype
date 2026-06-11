import { render, screen } from '@testing-library/react';
import App from './App';

jest.mock('@azure/msal-browser', () => ({
  PublicClientApplication: jest.fn().mockImplementation(() => ({
    loginRedirect: jest.fn(),
    logoutPopup: jest.fn(),
    acquireTokenSilent: jest.fn(),
  })),
}));

jest.mock('@azure/msal-react', () => ({
  MsalProvider: ({ children }) => children,
  useMsal: () => ({
    instance: {
      loginRedirect: jest.fn(),
      logoutPopup: jest.fn(),
      acquireTokenSilent: jest.fn(),
    },
    accounts: [],
  }),
}));

test('renders the sign in screen', () => {
  render(<App />);
  expect(screen.getByText(/permission manager/i)).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /sign in with microsoft/i })).toBeInTheDocument();
});
