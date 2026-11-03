OBJS = uvconvert.o
SRCS = uvconvert.cpp
LIBS = ../vnconv/libvnconv.a
INCLUDE_DIR = ../vnconv
OPTFLAGS="-O3"

all: uvconv
uvconv: ${OBJS} libvnconv
	${CXX} ${CXXFLAGS} ${OPTFLAGS} -o $@ ${OBJS} ${LIBS}

${OBJS}: %.o: %.cpp
	${CXX} ${CXXFLAGS} ${OPTFLAGS} -c $< -o $@ -I${INCLUDE_DIR}

.PHONY: libvnconv
libvnconv:
	cd ../vnconv && make all
depend:
	makedepend -- ${CFLAGS} -- ${SRCS} -I${INCLUDE_DIR}
	cd ../vnconv && make depend
clean:
	rm *.o
	cd ../vnconv && make clean

# DO NOT DELETE
