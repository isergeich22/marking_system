function errorHandler(err, req, res, next) {
    console.error(`[${new Date().toISOString()}] ${req.method} ${req.url} — ${err.message}`)
    const status = err.response?.status || err.status || 500
    res.status(status).json({ error: err.message || 'Internal server error' })
}

module.exports = errorHandler
