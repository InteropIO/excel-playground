import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, it, expect } from 'vitest';
import App from '../App';

describe('App', () => {
  it('renders the app title', () => {
    render(<App />);
    expect(screen.getByText('React App')).toBeInTheDocument();
  });

  it('renders initial count of 0', () => {
    render(<App />);
    expect(screen.getByText('0')).toBeInTheDocument();
    expect(screen.getByText('Count is 0')).toBeInTheDocument();
  });

  it('increments count when button is clicked', async () => {
    const user = userEvent.setup();
    render(<App />);
    
    const button = screen.getByRole('button', { name: /count is 0/i });
    await user.click(button);
    
    expect(screen.getByText('1')).toBeInTheDocument();
    expect(screen.getByText('Count is 1')).toBeInTheDocument();
  });

  it('increments count multiple times', async () => {
    const user = userEvent.setup();
    render(<App />);
    
    const button = screen.getByRole('button');
    
    await user.click(button);
    await user.click(button);
    await user.click(button);
    
    expect(screen.getByText('3')).toBeInTheDocument();
    expect(screen.getByText('Count is 3')).toBeInTheDocument();
  });
});