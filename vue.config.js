
chainWebpack: config => {
    // 修复HMR
    config.resolve.symlinks(true);

}