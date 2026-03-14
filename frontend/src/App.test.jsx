import { fireEvent, render, screen, waitFor } from '@testing-library/react'

import App from './App'

describe('App generator form', () => {
  it('sends preset and selected language in generate payload', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          output_path: '/tmp/generated.pptx',
          outline: [{ title: 't1', bullets: ['b1'] }],
          theme: { name: 'preset-tech', font_name: 'Segoe UI', title_size_pt: 38, body_size_pt: 20 },
        }),
      })
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({ path: '/tmp/generated.pptx', slide_count: 0, slides: [] }),
      })

    vi.stubGlobal('fetch', fetchMock)

    render(<App />)

    fireEvent.change(screen.getByLabelText('Topic / Prompt *'), {
      target: { value: 'AI Platform Plan' },
    })
    fireEvent.change(screen.getByTestId('preset-select'), {
      target: { value: 'Tech' },
    })
    fireEvent.change(screen.getByTestId('language-select'), {
      target: { value: 'ja-JP' },
    })

    fireEvent.click(screen.getByRole('button', { name: 'Generate PPT' }))

    await waitFor(() => expect(fetchMock).toHaveBeenCalledTimes(2))

    const generateRequest = JSON.parse(fetchMock.mock.calls[0][1].body)
    expect(generateRequest.topic).toBe('AI Platform Plan')
    expect(generateRequest.preset).toBe('Tech')
    expect(generateRequest.language).toBe('ja-JP')
  })
})
