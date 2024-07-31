FROM rust:1-bookworm AS build
RUN apt-get update && apt-get install -y git libclang-dev llvm-dev clang cmake zlib1g-dev
RUN rustup update stable
WORKDIR /project
COPY . /project
RUN cargo test --release
RUN cargo build --release

FROM debian:bookworm-slim
RUN apt-get update && apt-get install -y bash libbz2-1.0 libcurl4 liblzma5 && apt-get clean -y && rm -rf /var/lib/apt/lists/*
COPY --from=build /project/target/release/xlsxgenerator /usr/local/bin/
